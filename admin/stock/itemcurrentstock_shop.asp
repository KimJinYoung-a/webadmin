<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_summary.asp"-->
<!-- #include virtual="/lib/classes/stock/realjaegocls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%

const C_STOCK_DAY=7

dim itemgubun, itemid, itemoption, shopid, barcode
itemgubun  = request("itemgubun")
itemid     = request("itemid")
itemoption = request("itemoption")
barcode     = request("barcode")
shopid     = request("shopid")

if (barcode <> "") then
    if BF_IsMaybeTenBarcode(barcode) then
        itemgubun 	= BF_GetItemGubun(barcode)
    	itemid 		= BF_GetItemId(barcode)
    	itemoption 	= BF_GetItemOption(barcode)
    end if
else
    IF (itemid>=1000000) THEN
        barcode = itemgubun + "" + Format00(8,itemid) + "" + itemoption
    ELSE
        barcode = itemgubun + "" + Format00(6,itemid) + "" + itemoption
    END IF
end if


if (shopid = "") then
        shopid = "-"
end if

dim nowyyyymmdd
nowyyyymmdd = Left(now(), 10)


'==============================================================================
'��ǰ�⺻����
if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"

dim ojaegoitem
set ojaegoitem = new CRealJaeGo
ojaegoitem.FRectItemGubun = itemgubun
ojaegoitem.FRectItemID = itemid
ojaegoitem.FRectItemOption = itemoption
if itemid<>"" then
	ojaegoitem.GetOfflineItemDefaultData
end if

dim oitemoption
set oitemoption = new CItemOptionInfo
oitemoption.FRectItemID =  itemid
if itemid<>"" then
	oitemoption.getOptionList
end if

if (oitemoption.FResultCount<1) then
	itemoption = "0000"
end if


'==============================================================================
'��ǰ�������(current)
dim ocursummary
set ocursummary = new CShopItemSummary

ocursummary.FRectShopID =  shopid
ocursummary.FRectItemGubun =  itemgubun
ocursummary.FRectItemId =  itemid
ocursummary.FRectItemOption =  itemoption

if itemid<>"" then
	ocursummary.GetShopItemCurrentSummary
end if


'==============================================================================
'��ǰ�������(monthly)
dim omonsummary
set omonsummary = new CShopItemSummary

omonsummary.FRectShopID =  shopid
omonsummary.FRectItemGubun =  itemgubun
omonsummary.FRectItemId =  itemid
omonsummary.FRectItemOption =  itemoption

if itemid<>"" then
	omonsummary.GetShopItemMonthlySummaryList
end if


'==============================================================================
'��ǰ�������(last month)
dim olastmonsummary
set olastmonsummary = new CShopItemSummary

olastmonsummary.FRectShopID =  shopid
olastmonsummary.FRectItemGubun =  itemgubun
olastmonsummary.FRectItemId =  itemid
olastmonsummary.FRectItemOption =  itemoption

if itemid<>"" then
	olastmonsummary.GetShopItemLastMonthSummary
end if


'==============================================================================
'��ǰ�������(daily)
dim odaysummary
set odaysummary = new CShopItemSummary

odaysummary.FRectShopID =  shopid
odaysummary.FRectItemGubun =  itemgubun
odaysummary.FRectItemId =  itemid
odaysummary.FRectItemOption =  itemoption

if itemid<>"" then
	odaysummary.GetShopItemDailySummaryList
end if


dim i, buf
dim dstart, dend

%>

<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
}

function RefreshRecentStock(yyyymmdd,itemgubun,itemid,itemoption){
	if (confirm('�ֱ� 2�� ������ ���ΰ�ħ �Ͻðڽ��ϱ�?')){
		frmrefresh.mode.value="itemrecentipchulrefresh";
		frmrefresh.submit();
	}
}

function RefreshTodayStock(itemgubun,itemid,itemoption){
	if (confirm('���� ������ ���ΰ�ħ �Ͻðڽ��ϱ�?')){
		frmrefresh.mode.value="itemtodayipchulrefresh";
		frmrefresh.submit();
	}
}


function RefreshALLStock(yyyymmdd,itemgubun,itemid,itemoption){
	if (confirm('��ü ������ ���ΰ�ħ �Ͻðڽ��ϱ�?')){
		frmrefresh.mode.value="itemallipchulrefresh";
		frmrefresh.submit();
	}
}

function PopStockBaditem(fromdate,todate,itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poperritemlist.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popbaditemlist','width=800,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function popRealErrList(fromdate,todate,itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poperritemlist.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'poperritemlist','width=800,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemUpcheIpChulListOffLine(fromdate,todate,itemgubun,itemid,itemoption, ipchulflag, shopid){
	var popwin = window.open('/common/pop_upcheipgolist_off.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag + '&shopid=' + shopid,'poperritemlist','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemSellListOffLine(fromdate,todate,itemgubun,itemid,itemoption, ipchulflag, shopid){
	var popwin = window.open('/common/pop_selllist_off.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag + '&shopid=' + shopid,'poperritemlist','width=1000,height=600,scrollbars=yes,resizable=yes')
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
		<font color="red"><strong>OFFLINE�𺰻�ǰ�������Ȳ</strong></font></td>
		<td height="20" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			�������� ���� ��ǰ�� ��� �����Դϴ�.
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


<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method=get>
	<input type=hidden name=menupos value="<%= menupos %>">
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
        	���ڵ�: <input type=text name=barcode value="<%= barcode %>" size=14 maxlength=14>&nbsp;&nbsp;
			&nbsp;
        </td>
        <td valign="top" align="right">
        <a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- ǥ ��ܹ� ��-->

<% if ojaegoitem.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=<%= 5 + ojaegoitem.FResultCount -1 %> width="110" valign=top align=center><img src="<%= ojaegoitem.FItemList(0).FImageList %>" width="100" height="100"></td>
      	<td width="60"><b>*��ǰ����</b></td>
      	<td width="300">
      	<input type="button" value="����" onclick="PopItemSellEdit('<%= itemid %>');">
      	</td>
      	<td width="60">�ŷ���� :</td>
      	<td colspan=2><%= ojaegoitem.FItemList(0).getChargeDivName %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>��ǰ�ڵ� :</td>
      	<td><%= itemgubun %> <b><%= CHKIIF(ojaegoitem.FItemList(0).FItemID>=1000000,Format00(8,ojaegoitem.FItemList(0).FItemID),Format00(6,ojaegoitem.FItemList(0).FItemID)) %></b> <%= itemoption %></td>
      	<td>�Һ��ڰ� : </td>
      	<td colspan=2><%= FormatNumber(ojaegoitem.FItemList(0).FSellcash,0) %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>�귣��ID :</td>
      	<td><%= ojaegoitem.FItemList(0).FMakerid %></td>
      	<td>�����ް� : </td>
      	<td colspan=2><%= FormatNumber(ojaegoitem.FItemList(0).FBuycash,0) %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>��ǰ�� :</td>
      	<td><%= ojaegoitem.FItemList(0).FItemName %></td>
      	<td></td>
      	<td colspan=2></td>
    </tr>
    <% for i=0 to ojaegoitem.FResultCount -1 %>
	    <% if ojaegoitem.FItemList(i).Foptionusing<>"Y" then %>
	    <tr bgcolor="#FFFFFF">
	      	<td><font color="#AAAAAA">�ɼǸ� :</font></td>
	      	<td><font color="#AAAAAA"><%= ojaegoitem.FItemList(i).FItemOptionName %></font></td>
	      	<td></td>
	      	<td></td>
	      	<td></td>
	    </tr>
	    <% else %>

	    <% if ojaegoitem.FItemList(i).FItemOption=itemoption then %>
	    <tr bgcolor="#EEEEEE">
	    <% else %>
	    <tr bgcolor="#FFFFFF">
	    <% end if %>
	      	<td>�ɼǸ� :</td>
	      	<td><%= ojaegoitem.FItemList(i).FItemOptionName %></td>
	      	<td>�������� : </td>
	      	<td><%= ojaegoitem.FItemList(i).FLimitYn %> (<%= ojaegoitem.FItemList(i).GetLimitStr %>)</td>
	      	<td><%= ojaegoitem.FItemList(i).Fcurrno %></td>
	    </tr>
	    <% end if %>
    <% next %>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="20" valign="bottom">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<br>�ý��� ����� = �԰�/��ǰ�� + ��ü�԰�/��ǰ�� - ��OFF�Ǹ��� + ��Ÿ���/��ǰ��
        	<!--
	        <br>�ý��� ��ȿ��� = �ý��� ����� - �ҷ�
	        <br>�ǻ� ��� = �ý��� ��ȿ��� - �Է¿���

	        <br>����ľ� ��� = �ǻ� ��� - ON��ǰ�غ� - OFF��ǰ�غ�
		<br>��ü�ֹ� ��� = �ǻ� ��� - ON��ǰ�غ� - ON�����Ϸ� - OFF��ǰ�غ�
		-->
		<br><br><p>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->




<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="40" valign="bottom">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td><b>*�������</b></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->


<table width="100%" align="center" cellpadding="2" cellspacing="1" bgcolor="#BABABA" class="a">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="60">���԰�<br>(�ٹ�����)</td>
    	<td width="60">�ѹ�ǰ<br>(�ٹ�����)</td>
    	<td width="60">���԰�<br>(��ü)</td>
    	<td width="60">�ѹ�ǰ<br>(��ü)</td>
    	<td width="60">���Ǹ�</td>
    	<td width="60">�ѹ�ǰ</td>
    	<td width="60" bgcolor="F4F4F4">�ý������</td>
    	<td width="60">����</td>
    	<td width="60">�ҷ�</td>
    	<td width="60" bgcolor="F4F4F4">��ȿ���</td>
    	<td width="60">����</td>
    	<td width="60" bgcolor="F4F4F4">�������</td>
    	<td>���</td>
    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td><%= ocursummary.FOneItem.Flogicsipgono %></td>
    	<td><%= ocursummary.FOneItem.Flogicsreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fsellno %></td>
    	<td><%= ocursummary.FOneItem.Fresellno %></td>
    	<td bgcolor="F4F4F4"><b><%= ocursummary.FOneItem.Fsysstockno %></td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    </tr>
</table>












<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="40" valign="bottom">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td><b>*�Ϻ� ���⳻��</b></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->








<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="60">�Ͻ�</td>
    	<td width="60">�԰�<br>(�ٹ�����)</td>
    	<td width="60">��ǰ<br>(�ٹ�����)</td>
    	<td width="60">�԰�<br>(��ü)</td>
    	<td width="60">��ǰ<br>(��ü)</td>
    	<td width="60">�Ǹ�</td>
    	<td width="60">��ǰ</td>
    	<td width="60" bgcolor="F4F4F4">�ý������</td>
    	<td width="60">����</td>
    	<td width="60">�ҷ�</td>
    	<td width="60" bgcolor="F4F4F4">��ȿ���</td>
    	<td width="60">����</td>
    	<td width="60" bgcolor="F4F4F4">�������</td>
    	<td>���</td>
    </tr>
    <% for i=0 to omonsummary.FResultcount-1 %>
    <%
    dstart = omonsummary.FItemList(i).Fyyyymm + "-01"
    dend = Left(dateadd("m",1,dstart),7)+"-01"
    dend = Left(dateadd("d",-1,dend),10)
    %>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td><%= omonsummary.FItemList(i).Fyyyymm %></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Flogicsipgono %></a></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Flogicsreipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fbrandipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fbrandreipgono %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fsellno %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fresellno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= omonsummary.FItemList(i).Fsysstockno %></td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    </tr>
    <% next %>
    <% if (omonsummary.FResultcount < 1) then %>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td colspan="14" align="center">����Ÿ����</td>
    </tr>
    <% end if %>
    <%
    dstart = "2001-10-10"
    dend = Left(dateadd("m",-1,nowyyyymmdd),7)+"-01"
    dend = Left(dateadd("d",-1,dend),10)

    %>
    <tr bgcolor="#DDDDFF" height="25" align=center>
    	<td>�հ�<br>(2������)</td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Flogicsipgono %></a></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Flogicsreipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fbrandipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fbrandreipgono %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fsellno %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fresellno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= olastmonsummary.FOneItem.Fsysstockno %></td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    </tr>
    <% for i=0 to odaysummary.FResultcount-1 %>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td><%= odaysummary.FItemList(i).Fyyyymmdd %></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Flogicsipgono %></a></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Flogicsreipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fbrandipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fbrandreipgono %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fsellno %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fresellno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= odaysummary.FItemList(i).Fsysstockno %></td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    </tr>
    <% next %>
    <% if (odaysummary.FResultcount < 1) then %>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td colspan="14" align="center">����Ÿ����</td>
    </tr>
    <% end if %>
    <tr bgcolor="#DDDDFF" height="25" align=center>
    	<td>�հ�</td>
    	<td><%= ocursummary.FOneItem.Flogicsipgono %></td>
    	<td><%= ocursummary.FOneItem.Flogicsreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fsellno %></td>
    	<td><%= ocursummary.FOneItem.Fresellno %></td>
    	<td bgcolor="F4F4F4"><b><%= ocursummary.FOneItem.Fsysstockno %></td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    	<td bgcolor="F4F4F4"><b></b></td>
    	<td></td>
    </tr>
</table>







<% else %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#000000">
    <tr align="center" bgcolor="#DDDDFF">
    <td align=center bgcolor="#FFFFFF">�˻� ����� �����ϴ�.</td>
    </tr>
</table>
<% end if %>


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


<% if (oitemoption.FResultCount>0) and (itemoption="0000") then %>
<script language='javascript'>
alert('�ɼ� ���� �� �� �˻��ϼ���.');
</script>
<% elseif (oitemoption.FResultCount<1) and (itemoption<>"0000") then %>
<script language='javascript'>
alert('�� �˻��ϼ���.');
</script>
<% end if %>
<%
set oitemoption = Nothing
set ojaegoitem = Nothing
set ocursummary = Nothing
set omonsummary = Nothing
set ocursummary = Nothing
%>
<form name=frmrefresh method=post action="dostockrefresh.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="<%= itemoption %>">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->