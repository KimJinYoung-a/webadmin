<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->

<%

const C_STOCK_DAY=7

dim itemgubun, itemid, itemoption, shopid, barcode, makerid
itemgubun  = requestCheckVar(request("itemgubun"),2)
itemid     = requestCheckVar(request("itemid"),9)
itemoption = requestCheckVar(request("itemoption"),4)
barcode    = requestCheckVar(request("barcode"),32)
shopid     = requestCheckVar(request("shopid"),32)

'/����
if (C_IS_SHOP) then

	'//�������϶�
	if C_IS_OWN_SHOP then

		'/���α��� ���� �̸�
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/��ü
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")	'"7321"
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if

if (barcode <> "") then
    if Not (fnGetItemCodeByPublicBarcode(barcode,itemgubun,itemid,itemoption)) then
        if (Len(barcode)=12) then
            itemgubun   = Left(barcode, 2)
            itemid      = CStr(Mid(barcode, 3, 6) + 0)
            itemoption  = Right(barcode, 4)
        elseif (Len(barcode)=14) then
            itemgubun   = Left(barcode, 2)
            itemid      = CStr(Mid(barcode, 3, 8) + 0)
            itemoption  = Right(barcode, 4)
        end if
    end if
elseif (itemid<>"") then
    if (itemid>=1000000) then
        barcode = itemgubun + "" + Format00(8,itemid) + "" + itemoption
    else
        barcode = itemgubun + "" + Format00(6,itemid) + "" + itemoption
    end if
end if


if (shopid = "") then
        shopid = ""
end if

dim nowyyyymmdd
nowyyyymmdd = Left(now(), 10)


'==============================================================================
'��ǰ�⺻����
if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"

dim ojaegoitem
set ojaegoitem = new COffShopItem
ojaegoitem.FRectItemGubun   = itemgubun
ojaegoitem.FRectItemID      = itemid
ojaegoitem.FRectItemOption  = itemoption
ojaegoitem.FRectShopid      = shopid
if (itemid<>"") then
	ojaegoitem.GetOffOneItem
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


dim BasicMonth
BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)

'==============================================================================
'��ǰ�������(daily)
dim odaysummary
set odaysummary = new CShopItemSummary

odaysummary.FRectShopID =  shopid
odaysummary.FRectItemGubun =  itemgubun
odaysummary.FRectItemId =  itemid
odaysummary.FRectItemOption =  itemoption
odaysummary.FRectStartDate  =  BasicMonth + "-01"
if itemid<>"" then
	odaysummary.GetShopItemDailySummaryList
end if


dim i, buf
dim dstart, dend

dim sysstockSum
dim availstockSum
dim realstockSum

sysstockSum    =0
availstockSum  =0
realstockSum   =0

dim IsUpcheWitakItem
if (ojaegoitem.FResultCount>0) then
    IsUpcheWitakItem = (ojaegoitem.FOneItem.Fcomm_cd="B012")
else
    IsUpcheWitakItem = False
end if

%>

<script type='text/javascript'>

function popOffItemEdit(ibarcode){
	<% if (C_IS_SHOP) then %>

		//�������϶�
		<% if C_IS_OWN_SHOP then %>
			var popwin = window.open('/admin/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
			popwin.focus();
		<% else %>
			return;
		<% end if %>
	<% else %>
		<% if (C_IS_Maker_Upche) then %>
			var popwin = window.open('/admin/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
			popwin.focus();
		<% else %>
			<% if (Not C_ADMIN_USER) then %>
				return;
			<% else %>
				var popwin = window.open('/admin/offshop/popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
				popwin.focus();
			<% end if %>
		<% end if %>
	<% end if %>
}

function refreshOffStockByItem(itemgubun,itemid,itemoption){
    if (frmrefresh.shopid.value.length<1){
        alert('������ ������ �˻� �� ����ϼ���.');
        return;
    }

    if (confirm('��� ������ ��ü ���ΰ�ħ �Ͻðڽ��ϱ�?')){
		frmrefresh.mode.value="OFFitemAllRefresh";
		frmrefresh.submit();
	}
}

function refreshOffStockByItemV2(itemgubun,itemid,itemoption){
    if (frmrefresh.shopid.value.length<1){
        alert('������ ������ �˻� �� ����ϼ���.');
        return;
    }

    if (confirm('��� ������ ��ü ���ΰ�ħ �Ͻðڽ��ϱ�?')){
		frmrefresh.mode.value="OFFStockitemRecentRefresh";
		frmrefresh.submit();
	}
}

function refreshAccStockShop(comp,yyyymm){
	var frm =document.frmrefresh;
	frm.mode.value = "itemAccStockShop";
	frm.yyyymm.value = yyyymm;
	

    var confirmstr = yyyymm+'�� ������ü �⸻��� ���ΰ�ħ �Ͻðڽ��ϱ�?'

    if (confirm(confirmstr)){
		comp.disabled=true;
		frm.submit();
	}
}




function popOffErrInput(shopid,itemgubun,itemid,itemoption){
    //�Է�â���� üũ : ��ü��Ź ��ǰ.
    <% if (C_IS_Maker_Upche) and (Not IsUpcheWitakItem) then %>
    alert('������ �����ϴ�. - ��ü��Ź ��ǰ�� ��� ���� ����.');
    return;
    <% else %>
    var popwin = window.open('/common/offshop/popOffrealerrinput.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popOffrealerrinput','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
	<% end if %>
}

function popOffStockBaditem(fromdate,todate,itembarcode,errType,shopid){
    <% if (C_ADMIN_USER) then %>
	var popwin = window.open('/admin/stock/off_baditem_list.asp?fromdate=' + fromdate + '&todate=' + todate + '&itembarcode=' + itembarcode +  '&errType=' + errType + '&shopid=' + shopid,'popoffbaditemlist','width=900,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
	<% end if %>
}

function PopItemUpcheIpChulListOffLine(fromdate,todate,itemgubun,itemid,itemoption, ipchulflag, shopid){
    <% if (C_ADMIN_USER) then %>
	var popwin = window.open('/common/pop_upcheipgolist_off.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag + '&shopid=' + shopid,'poperritemlist','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
	<% end if %>
}

function PopItemSellListOffLine(fromdate,todate,itemgubun,itemid,itemoption, ipchulflag, shopid){
    <% if (C_ADMIN_USER) then %>
	var popwin = window.open('/common/pop_selllist_off.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag + '&shopid=' + shopid,'poperritemlist','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
	<% end if %>
}


function popAsgnAccGbn(yyyymm,shopid,itemgubun,itemid,itemoption){
return;
    <% if (C_ADMIN_AUTH) then %>
    var popwin = window.open('/admin/newreport/popAssignMonthlyAccMwgubun.asp?stockPlace=S&yyyymm='+yyyymm+'&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&shopid=' + shopid,'poperritemlist','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
    <% end if %>
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
			<% if (C_IS_SHOP) then %>
				<% if C_IS_OWN_SHOP then %>
					���� : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %> &nbsp;&nbsp;
				<% else %>
					���� : <%= shopid %>
				<% end if %>
			<% else %>
				<% if (C_IS_Maker_Upche) then %>
					���� : <% drawSelectBoxOpenOffShop "shopid",shopid %>
				<% else %>
					<% if (Not C_ADMIN_USER) then %>
					<% else %>
						���� : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %> &nbsp;&nbsp;
					<% end if %>
				<% end if %>
			<% end if %>

        	<% if (C_IS_Maker_Upche) then %>
        	    <input type="hidden" name="barcode" value="<%= barcode %>">
        	<% else %>
        	���ڵ�: <input type="text" class="text" name="barcode" value="<%= barcode %>" size=16 maxlength=20 <%= ChkIIF(C_ADMIN_USER,"","readonly") %> >&nbsp;&nbsp;
			&nbsp;
			<% end if %>
			<input type="button" class="button" value=" �� �� " onClick="document.frm.submit();">
        </td>
        <td valign="top" align="right">
            <% if (C_ADMIN_USER) or (C_OFF_AUTH) then %>
            <input type="button" class="button" value="��� ���� ��ħ" onClick="refreshOffStockByItem('<%= itemgubun %>','<%= itemid %>','<%= itemoption %>')">
			&nbsp;
			<input type="button" class="button" value="��� ���� ��ħ V2" onClick="refreshOffStockByItemV2('<%= itemgubun %>','<%= itemid %>','<%= itemoption %>')">
            <% end if %>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- ǥ ��ܹ� ��-->

<% if ojaegoitem.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
    	<td rowspan="5" width="110" valign=top align=center><img src="<%= ojaegoitem.FOneItem.GetImageList %>" width="100" height="100"></td>
      	<td width="60"><b>*��ǰ����</b></td>
      	<td width="300">
			<% if (C_IS_SHOP) then %>
				<% if C_IS_OWN_SHOP then %>
					<input type="button" class="button" value="����" onclick="popOffItemEdit('<%= barcode %>');">
				<% else %>
					
				<% end if %>
			<% else %>
				<% if (C_IS_Maker_Upche) then %>
					<input type="button" class="button" value="����" onclick="popOffItemEdit('<%= barcode %>');">
				<% else %>
					<% if (Not C_ADMIN_USER) then %>
					<% else %>
						<input type="button" class="button" value="����" onclick="popOffItemEdit('<%= barcode %>');">
					<% end if %>
				<% end if %>
			<% end if %>
      	</td>
      	<td width="80">�ŷ���� </td>
      	<td colspan=2>
          	<% if Not ojaegoitem.FOneItem.IsShopContractExists then %>
          	<font color="red"><strong>������</strong></font>
          	<% else %>
          	<%= GetJungsanGubunName(ojaegoitem.FOneItem.FComm_cd) %>
          	    <% if (C_ADMIN_USER) then %>
          	    [<%= ojaegoitem.FOneItem.FMakerMargin %> -&gt; <%= ojaegoitem.FOneItem.FshopMargin %>]
          	    <% end if %>
          	<% end if %>
      	</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>��ǰ�ڵ�</td>
      	<td><%= ojaegoitem.FOneItem.GetBarCode %></td>
      	<td>�ǸŰ�</td>
      	<td colspan=2>
      	    <% if (ojaegoitem.FOneItem.IsOffSaleItem) then %>
      	    <strike><%= FormatNumber(ojaegoitem.FOneItem.FShopItemOrgprice,0) %></strike>
      	    &nbsp;&nbsp;
      	    <%= FormatNumber(ojaegoitem.FOneItem.Fshopitemprice,0) %>
      	    <% else %>
      	    <%= FormatNumber(ojaegoitem.FOneItem.Fshopitemprice,0) %>
      	    <% end if %>
      	</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>�귣��ID</td>
      	<td><%= ojaegoitem.FOneItem.FMakerid %></td>
      	<% if (C_IS_Maker_Upche) or (C_ADMIN_USER) then %>
      	<td>���԰�(��ü)</td>
      	<td colspan=2>
          	<% if ojaegoitem.FOneItem.IsShopContractExists then %>
          	    <%= FormatNumber(ojaegoitem.FOneItem.GetOfflineBuycash,0) %>
          	<% end if %>
      	</td>
      	<% elseif (C_IS_SHOP) then %>
      	<td>���ް�(SHOP)</td>
      	<td colspan=2>
      	<% if ojaegoitem.FOneItem.IsShopContractExists then %>
      	    <%= FormatNumber(ojaegoitem.FOneItem.GetOfflineSuplycash,0) %>
      	<% end if %>
        </td>
      	<% else %>
      	<td></td>
      	<td colspan=2></td>
      	<% end if %>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>��ǰ��</td>
      	<td>
      	    <%= ojaegoitem.FOneItem.FShopItemName %>
      	    <% if (ojaegoitem.FOneItem.FShopItemOptionName<>"") then %>
      	    <font color="blue">[<%= ojaegoitem.FOneItem.FShopItemOptionName %>]</font>
      	    <% end if %>
      	</td>
      	<% if (C_ADMIN_USER) then %>
      	<td>���ް�(SHOP)</td>
      	<td colspan=2>
      	<% if ojaegoitem.FOneItem.IsShopContractExists then %>
      	    <%= FormatNumber(ojaegoitem.FOneItem.GetOfflineSuplycash,0) %>
      	<% end if %>
        </td>
        <% else %>
        <td></td>
      	<td colspan=2></td>
        <% end if %>
    </tr>
</table>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="20" valign="bottom">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<br>�ý��� ���     = �԰�/��ǰ�� + ��ü�԰�/��ǰ�� + ���Ǹ�/��ǰ
	        <br>�ǻ� ���       = �ý������ + �Է¿���
	        <br>��ȿ���        = �ý��� ��� + �Է¿��� + ���� <!-- + �ҷ� -->

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
        <td width="60">&nbsp;</td>
    	<td width="60">���԰�<br>(�ٹ�����)</td>
    	<td width="60">�ѹ�ǰ<br>(�ٹ�����)</td>
    	<td width="60">���԰�<br>(��ü)</td>
    	<td width="60">�ѹ�ǰ<br>(��ü)</td>
    	<td width="60">���Ǹ�</td>
    	<td width="60">�ѹ�ǰ</td>
    	<td width="60" bgcolor="F4F4F4">�ý������<br>(����)</td>
    	<td width="60">����</td>
    	<td width="60" bgcolor="F4F4F4">�ǻ����<br>(����)</td>
    	<td width="60">����</td>
    	<!-- <td width="60">�ҷ�</td> -->
    	<td width="60" bgcolor="F4F4F4">��ȿ���<br>(����)</td>
		<td width="60">�����</td>
		<td width="60">��ǰ��</td>
		<td width="60" bgcolor="F4F4F4">�������<br>(����)</td>
    	<td>���</td>
    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
        <td>&nbsp;</td>
    	<td><%= ocursummary.FOneItem.Flogicsipgono %></td>
    	<td><%= ocursummary.FOneItem.Flogicsreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fsellno %></td>
    	<td><%= ocursummary.FOneItem.Fresellno %></td>
    	<td bgcolor="F4F4F4"><b><%= ocursummary.FOneItem.Fsysstockno %></b></td>
    	<td><%= ocursummary.FOneItem.Ferrrealcheckno %></td>
    	<td><b><%= ocursummary.FOneItem.Frealstockno %></b></td>
    	<td><%= ocursummary.FOneItem.Ferrsampleitemno %></td>
    	<!-- <td><%= ocursummary.FOneItem.Ferrbaditemno %></td> -->
    	<td bgcolor="F4F4F4"><%= ocursummary.FOneItem.getAvailStock %></td>
		<td bgcolor="F4F4F4"><%= ocursummary.FOneItem.Flogischulgo %></td>
		<td bgcolor="F4F4F4"><%= ocursummary.FOneItem.Flogisreturn %></td>
		<td bgcolor="FFDDDD"><%= ocursummary.FOneItem.getShopRealStock %></td>
    	<td>
    	    <% if ocursummary.FOneItem.Fpreorderno>0 then %>
    	    ���ֹ� : <%= ocursummary.FOneItem.Fpreorderno %>
    	        <% if (ocursummary.FOneItem.Fpreorderno<>ocursummary.FOneItem.FpreordernoFix) then %>
                    => <strong> <%= ocursummary.FOneItem.FpreordernoFix %></strong>
    	        <% end if %>
    	    <% end if %>
    	</td>
    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
        <td colspan="9"></td>
        <td></td>
        <td><input type="button" class="button" value="����" onClick="popOffErrInput('<%= shopid %>','<%= itemgubun %>','<%= itemid %>','<%= itemoption %>');"></td>
        <!-- <td><input type="button" class="button" value="�ҷ�"></td> -->
        <td></td>
        <td></td>
		<td></td>
		<td bgcolor="#FFDDDD"><input type="button" class="button" value="�ǻ�" onClick="popOffErrInput('<%= shopid %>','<%= itemgubun %>','<%= itemid %>','<%= itemoption %>');"></td>
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
    	<td width="60" bgcolor="F4F4F4">�ý������<br>(����)</td>
    	<td width="60">����</td>
    	<td width="60" bgcolor="F4F4F4">�ǻ����<br>(����)</td>
    	<td width="60">����</td>
    	<!-- <td width="60">�ҷ�</td> -->
    	<td width="60" bgcolor="F4F4F4">��ȿ���<br>(����)</td>
    	<td>���</td>
    </tr>
    <% for i=0 to omonsummary.FResultcount-1 %>
    <%
    dstart = omonsummary.FItemList(i).Fyyyymm + "-01"
    dend = Left(dateadd("m",1,dstart),7)+"-01"
    dend = Left(dateadd("d",-1,dend),10)

    sysstockSum    = sysstockSum    + omonsummary.FItemList(i).Fsysstockno
    availstockSum  = availstockSum  + omonsummary.FItemList(i).getAvailStock
    realstockSum   = realstockSum   + omonsummary.FItemList(i).Frealstockno
    %>
    <tr bgcolor="#FFFFFF" height="20" align=center>
    	<td><%= omonsummary.FItemList(i).Fyyyymm %></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Flogicsipgono %></a></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Flogicsreipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fbrandipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fbrandreipgono %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fsellno %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= omonsummary.FItemList(i).Fresellno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= sysstockSum %></b></td>
    	<td><a href="javascript:popOffStockBaditem('<%= dstart %>','<%= dend %>','<%= ojaegoitem.FOneItem.GetBarCode %>','D','<%= shopid %>')"><%= omonsummary.FItemList(i).Ferrrealcheckno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= realstockSum %></b></td>
    	<td><a href="javascript:popOffStockBaditem('<%= dstart %>','<%= dend %>','<%= ojaegoitem.FOneItem.GetBarCode %>','S','<%= shopid %>')"><%= omonsummary.FItemList(i).Ferrsampleitemno %></a></td>
    	<!-- <td><a href="javascript:popOffStockBaditem('<%= dstart %>','<%= dend %>','<%= ojaegoitem.FOneItem.GetBarCode %>','B','<%= shopid %>')"><%= omonsummary.FItemList(i).Ferrbaditemno %></a></td> -->
    	<td bgcolor="F4F4F4"><b><%= availstockSum %></b></td>
    	<td>
    	    <a href="javascript:popAsgnAccGbn('<%=omonsummary.FItemList(i).Fyyyymm%>','<%=shopid%>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>');">
    	   <%=omonsummary.FItemList(i).Fcomm_cd%>
    	    /
    	    <%=omonsummary.FItemList(i).FCenterMwdiv%>
    	    /
    	    <% if Not isNULL(omonsummary.FItemList(i).FAccSysstockno)  THEN %>
    	       <% if sysstockSum<>omonsummary.FItemList(i).FAccSysstockno then %>
    	            <font color=red><%=omonsummary.FItemList(i).FAccSysstockno%></font>
    	       <% else %>
    	            <%=omonsummary.FItemList(i).FAccSysstockno%>
    	        <% end if %>
    	    <% end if %>
    	   /
    	    </a>
    	</td>
    </tr>
    <% next %>
    <% if (omonsummary.FResultcount < 1) then %>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td colspan="14" align="center">[���� ����Ÿ ����]</td>
    </tr>
    <% end if %>
    <%
    dstart = "2001-10-10"
    dend = Left(dateadd("m",-1,nowyyyymmdd),7)+"-01"
    dend = Left(dateadd("d",-1,dend),10)

    %>
    <tr bgcolor="#DDDDFF" height="20" align=center>
    	<td>�հ�<br>(2������)</td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Flogicsipgono %></a></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Flogicsreipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fbrandipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fbrandreipgono %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fsellno %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= dstart %>','<%= dend %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= olastmonsummary.FOneItem.Fresellno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= sysstockSum %></b><%= ChkIIF(sysstockSum<>olastmonsummary.FOneItem.Fsysstockno,"<font color=red>(" & (olastmonsummary.FOneItem.Fsysstockno) & ")</font>","") %></td>
    	<td><a href="javascript:popOffStockBaditem('<%= dstart %>','<%= dend %>','<%= ojaegoitem.FOneItem.GetBarCode %>','D','<%= shopid %>')"><%= olastmonsummary.FOneItem.Ferrrealcheckno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= olastmonsummary.FOneItem.Frealstockno %></b><%= ChkIIF(realstockSum<>olastmonsummary.FOneItem.Frealstockno,"<font color=red>(" + CStr(realstockSum) + ")</font>","") %></td>
    	<td><a href="javascript:popOffStockBaditem('<%= dstart %>','<%= dend %>','<%= ojaegoitem.FOneItem.GetBarCode %>','S','<%= shopid %>')"><%= olastmonsummary.FOneItem.Ferrsampleitemno %></a></td>
    	<!-- <td><a href="javascript:popOffStockBaditem('<%= dstart %>','<%= dend %>','<%= ojaegoitem.FOneItem.GetBarCode %>','B','<%= shopid %>')"><%= olastmonsummary.FOneItem.Ferrbaditemno %></a></td> -->
    	<td bgcolor="F4F4F4"><b><%= availstockSum %></b><%= ChkIIF(availstockSum<>olastmonsummary.FOneItem.getAvailStock,"<font color=red>(" & (olastmonsummary.FOneItem.getAvailStock) & ")</font>","") %></td>
    	<td></td>
    </tr>
    <% for i=0 to odaysummary.FResultcount-1 %>
    <%
    sysstockSum    = sysstockSum    + odaysummary.FItemList(i).Fsysstockno
    availstockSum  = availstockSum  + odaysummary.FItemList(i).getAvailStock
    realstockSum   = realstockSum   + odaysummary.FItemList(i).Frealstockno
    %>
    <tr bgcolor="#FFFFFF" height="20" align=center>
    	<td><%= odaysummary.FItemList(i).Fyyyymmdd %></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Flogicsipgono %></a></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Flogicsreipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fbrandipgono %></a></td>
    	<td><a href="javascript:PopItemUpcheIpChulListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fbrandreipgono %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fsellno %></a></td>
    	<td><a href="javascript:PopItemSellListOffLine('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= shopid %>');"><%= odaysummary.FItemList(i).Fresellno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= sysstockSum %></b></td>
    	<td><a href="javascript:popOffStockBaditem('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= ojaegoitem.FOneItem.GetBarCode %>','D','<%= shopid %>')"><%= odaysummary.FItemList(i).Ferrrealcheckno %></a></td>
    	<td bgcolor="F4F4F4"><b><%= realstockSum %></b></td>
    	<td><a href="javascript:popOffStockBaditem('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= ojaegoitem.FOneItem.GetBarCode %>','S','<%= shopid %>')"><%= odaysummary.FItemList(i).Ferrsampleitemno %></a></td>
    	<!-- <td><a href="javascript:popOffStockBaditem('<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= odaysummary.FItemList(i).Fyyyymmdd %>','<%= ojaegoitem.FOneItem.GetBarCode %>','B','<%= shopid %>')"><%= odaysummary.FItemList(i).Ferrbaditemno %></a></td> -->
    	<td bgcolor="F4F4F4"><b><%= availstockSum %></b></td>
    	<td></td>
    </tr>
    <% next %>
    <% if (odaysummary.FResultcount < 1) then %>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td colspan="14" align="center">[�Ϻ� ���Ӹ� ����Ÿ�� �������� �ʽ��ϴ�.]</td>
    </tr>
    <% end if %>
    <tr bgcolor="#DDDDFF" height="20" align=center>
    	<td>�հ�</td>
    	<td><%= ocursummary.FOneItem.Flogicsipgono %></td>
    	<td><%= ocursummary.FOneItem.Flogicsreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandipgono %></td>
    	<td><%= ocursummary.FOneItem.Fbrandreipgono %></td>
    	<td><%= ocursummary.FOneItem.Fsellno %></td>
    	<td><%= ocursummary.FOneItem.Fresellno %></td>
    	<td bgcolor="F4F4F4"><b><%= sysstockSum %></b></td>
    	<td><%= ocursummary.FOneItem.Ferrrealcheckno %></td>
    	<td bgcolor="F4F4F4"><b><%= realstockSum %></b></td>
    	<td><%= ocursummary.FOneItem.Ferrsampleitemno %></td>
    	<!-- <td><%= ocursummary.FOneItem.Ferrbaditemno %></td> -->
    	<td bgcolor="F4F4F4"><b><%= availstockSum %></b></td>
    	<td>
		<% if (C_ADMIN_USER) or (C_OFF_AUTH) then %>
			<input type="button" value="�⸻���ۼ� <%=LEFT(dateadd("m",-1,now()),7)%>" onClick="refreshAccStockShop(this,'<%=LEFT(dateadd("m",-1,now()),7)%>')">
			<input type="button" value="�⸻���ۼ� <%=LEFT(dateadd("m",-0,now()),7)%>" onClick="refreshAccStockShop(this,'<%=LEFT(dateadd("m",-0,now()),7)%>')">
		<% end if %>
		</td>
    </tr>
</table>

<% else %>
<table width="100%" height="30" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
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



<%
set ojaegoitem = Nothing
set ocursummary = Nothing
set omonsummary = Nothing
set ocursummary = Nothing
%>
<form name=frmrefresh method=post action="/common/offshop/shop_stockrefresh_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="<%= itemoption %>">
<input type="hidden" name="yyyymm" value="">
</form>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
