<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itembarcodecls.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim itembarcode
dim itemgubun,itemid,itemoption
dim BasicMonth

itembarcode = requestCheckVar(request("itembarcode"),20)
itemgubun 	= requestCheckVar(request("itemgubun"),2)
itemid		 = requestCheckVar(request("itemid"),10)
itemoption	 = requestCheckVar(request("itemoption"),4)
BasicMonth	 = requestCheckVar(request("BasicMonth"),10)

if (Len(itembarcode)=12) then
	itemgubun 	= left(itembarcode,2)
	itemid		= CLng(mid(itembarcode,3,6))
	if (itemoption="") then itemoption = right(itembarcode,4)
	itembarcode = itemgubun + Format00(6,itemid) + itemoption
elseif (Len(itembarcode)=14) then
	itemgubun 	= left(itembarcode,2)
	itemid		= CLng(mid(itembarcode,3,8))
	if (itemoption="") then itemoption = right(itembarcode,4)
	itembarcode = itemgubun + Format00(8,itemid) + itemoption
elseif (Len(itembarcode)<>0) and (itemid<>"") then
	if itemgubun="" then itemgubun = "10"
	itemid = itembarcode
	if (itemoption="") then itemoption  = "0000"
elseif (Len(itembarcode)>7) then
    '''���ڵ��ΰ�� �˻��� ��ǰ�ڵ� ������.
    call fnGetItemCodeByPublicBarcode(itembarcode, itemgubun, itemid, itemoption)
else
    if itemgubun="" then itemgubun = "10"
    if (itemid="") then itemid = itembarcode
    if (itemoption="") then itemoption  = "0000"

    if (itemid>=1000000) then
        itembarcode = itemgubun + Format00(8,itemid) + itemoption
    else
        itembarcode = itemgubun + Format00(6,itemid) + itemoption
    end if
end if



dim oitem
set oitem = new CItem
oitem.FRectItemID = itemid

if (itemid<>"") and (itemgubun="10") then
	oitem.GetOneItem
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if (itemid<>"")  and (itemgubun="10") then
	oitemoption.GetItemOptionInfo
end if

''������ǰ
dim ooffitem
set ooffitem = new COffShopItem
ooffitem.FRectItemGubun = itemgubun
ooffitem.FRectItemID    = itemid
ooffitem.FRectItemOption = itemoption
if (itemgubun<>"") and (itemid<>"") and (itemoption<>"") and (itemgubun<>"10") then
	ooffitem.GetOffOneItem
end if

dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectStartDate = BasicMonth + "-01"
osummarystock.FRectItemGubun = itemgubun
osummarystock.FRectItemID    =  itemid
osummarystock.FRectItemOption =  itemoption
if itemid<>"" then
	osummarystock.GetCurrentItemStock
end if


dim otodayerritem
set otodayerritem = new CSummaryItemStock
otodayerritem.FRectItemGubun = itemgubun
otodayerritem.FRectItemID =  itemid
otodayerritem.FRectItemOption =  itemoption
if itemid<>"" then
    otodayerritem.GetTodayErrItem
end if

dim difftime
if (osummarystock.FResultcount>0) then
    difftime = ABS(datediff("h",osummarystock.FOneItem.Flastupdate,now()))
end if

dim i
dim IsVaildCode, IsStockExists
IsVaildCode = False
if (oitemoption.FResultCount>0) then
    for i=0 to oitemoption.FResultCount-1
        if (oitemoption.FITemList(i).FItemOption=itemoption) then
            IsVaildCode = (oitem.FResultCount>0)
            exit For
        end if
    next
else
    IsVaildCode = ((oitem.FResultCount>0) and (itemoption="0000")) or (ooffitem.FResultCount>0)
end if

IsStockExists = (osummarystock.FResultCount>0)
%>
<script language='javascript'>


function RecalcuErr(){
	var checkstock = calcufrm.checkstock.value;  // ����ľ����.

	calcufrm.todayerrrealcheckno.value = checkstock-calcufrm.orgrealstock.value - calcufrm.todaybaljuno.value;
	calcufrm.errrealcheckno.value = checkstock - calcufrm.availsysstock.value - calcufrm.todaybaljuno.value;
}

function SaveErr(){
//	if (<%= difftime %>>=4){
//		alert('���� ������Ʈ�ð��� 4�ð� ���� �Դϴ�. \n���� ���ΰ�ħ�� ����ϼ���.');
//		return;
//	}

	var realstock = calcufrm.checkstock.value;
	if (!IsInteger(realstock)){
		alert('���ڸ� �Է��ϼ���.');
		calcufrm.checkstock.focus();
		return;
	}

	if (confirm('�ǻ������ �����Ͻðڽ��ϱ�?')){
		frmrefresh.mode.value ="errcheckupdate";
		frmrefresh.realstock.value = realstock;
		frmrefresh.submit();
	}
}

function GetOnLoad(){
	<% if Not IsVaildCode then %>
	alert('��ǰ�ڵ尡 ��Ȯ���� �ʽ��ϴ�. ��˻� �ϼ���.');
	document.frm.itembarcode.select();
	document.frm.itembarcode.focus();
	<% else %>
	if (calcufrm.checkstock){
	    document.calcufrm.checkstock.select();
	    document.calcufrm.checkstock.focus();
	}
	<% end if %>
}
window.onload=GetOnLoad;

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get>
	<input type=hidden name=BasicMonth value="<%= BasicMonth %>">
	<!-- ��ܹ� ���� -->
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="2">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
				<tr>
					<td>
						<img src="/images/icon_arrow_down.gif" align="absbottom">
				        <font color="red">&nbsp;<strong>���(����)�Է�</strong></font>
				    </td>
				    <td align="right">
						��ǰ�ڵ�:
						<input type=text class="text" name=itembarcode value="<%= itembarcode %>" size=16 maxlength=16 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ document.frm.submit(); return false;}">
						<!--
			        	<input type=text class="text_ro" name=itemgubun value="<%= itemgubun %>" size=2 maxlength=2 readonly>
			        	<input type=text class="text" name=itemid value="<%= itemid %>" size=9 maxlength=9>
			        	<input type="text" class="text_ro" value="<%= itemoption %>" size=4 maxlength=4 readonly>
			        	-->
						&nbsp;

						<% if oitemoption.FResultCount>0 then %>

						<select class="select" name="itemoption">
						<option value="0000">----
						<% for i=0 to oitemoption.FResultCount-1 %>
						<option value="<%= oitemoption.FITemList(i).FItemOption %>" <% if itemoption=oitemoption.FITemList(i).FItemOption then response.write "selected" %> ><%= oitemoption.FITemList(i).FOptionName %>
						<% next %>
						</select>
						<% end if %>

        				<input type="button" class="button" value="�˻�" onclick="document.frm.submit();">
        				<!-- �ֱ� ���� ���ΰ�ħ �� �Էµ�
				        <%= BasicMonth %>-01 ~
				        <input type="button" value="���ΰ�ħ" onclick="RefreshRecentStock();">
				        -->
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<!-- ��ܹ� �� -->
	</form>
</table>

<p>

<% if (oitem.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=<%= 6 + oitemoption.FResultCount -1 %> width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
      	<td width="60">��ǰ�ڵ�</td>
      	<td width="300">
      		10<Strong><%= CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) %></Strong><%= itemoption %>
      	</td>
      	<td width="60"></td>
      	<td colspan=2></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>�귣��ID</td>
      	<td><%= oitem.FOneItem.FMakerid %></td>
      	<td>�Ǹſ���</td>
      	<td colspan=2><%= fnColor(oitem.FOneItem.FSellyn,"yn") %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>��ǰ��</td>
      	<td><%= oitem.FOneItem.FItemName %></td>
      	<td>��뿩��</td>
      	<td colspan=2><%= fnColor(oitem.FOneItem.FIsUsing,"yn") %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>�ǸŰ�</td>
      	<td>
      		<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
      		&nbsp;&nbsp;
      		<font color="<%= mwdivColor(oitem.FOneItem.FMwDiv) %>"><%= oitem.FOneItem.getMwDivName %></font>
      	    <% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
			<% end if %>
			&nbsp;&nbsp;
			<!-- ���ο���/�������뿩�� -->
			<% if (oitem.FOneItem.FSailYn="Y") then %>
			    <font color=red>
			    <% if (oitem.FOneItem.Forgprice<>0) then %>
			        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %> %
			    <% end if %>
			     ����
			    </font>
			<% end if %>

			<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>

			    <font color=green><%= oitem.FOneItem.GetCouponDiscountStr %> ����
			    (<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>)</font>
			<% end if %>

      	</td>
      	<td>��������</td>
      	<td colspan=2>
      		<% if oitem.FOneItem.Fdanjongyn="Y" then %>
			<font color="#33CC33">����</font>
			<% elseif oitem.FOneItem.Fdanjongyn="M" then %>
			<font color="#CC3333">MDǰ��</font>
			<% elseif oitem.FOneItem.Fdanjongyn="S" then %>
			<font color="#33CC33">�Ͻ�ǰ��</font>
			<% else %>
			������
			<% end if %>
		</td>
    </tr>
     <% if oitemoption.FResultCount>1 then %>
	    <!-- �ɼ��� �ִ°�� -->
	    <% for i=0 to oitemoption.FResultCount -1 %>
	    	<% if oitemoption.FITemList(i).Fitemoption=itemoption then %>
	    	<tr bgcolor="#FFFFFF">
	    		<td>�ɼǸ�</td>
		      	<td><%= oitemoption.FITemList(i).FOptionName %> (<%= fnColor(oitemoption.FITemList(i).FOptIsUsing,"yn") %>)</td>
		      	<td>��������</td>
		      	<td><%= fnColor(oitem.FOneItem.Flimityn,"yn") %> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
		      	<td>���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
		    </tr>
		    <% end if %>
		<% next %>
	<% else %>
    	<tr bgcolor="#FFFFFF">
	      	<td>�ɼǸ�</td>
	      	<td>-</td>
	      	<td>��������</td>
	      	<td><%= fnColor(oitem.FOneItem.Flimityn,"yn") %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
	      	<td>���� ����� (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)</td>
	    </tr>
    <% end if %>

</table>

<% elseif (ooffitem.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=5 width="110" valign=top align=center><img src="<%= ooffitem.FOneItem.FOffimgList %>" width="100" height="100"></td>
      	<td width="60">��ǰ�ڵ�</td>
      	<td width="300">
      		<%= ooffitem.FOneItem.GetBarCode %>
      	</td>
      	<td width="60"></td>
      	<td colspan=2></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>�귣��ID</td>
      	<td><%= ooffitem.FOneItem.FMakerid %></td>
      	<td></td>
      	<td colspan=2></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>��ǰ��</td>
      	<td><%= ooffitem.FOneItem.Fshopitemname %></td>
      	<td>��뿩��</td>
      	<td colspan=2><%= fnColor(ooffitem.FOneItem.FIsUsing,"yn") %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>�ǸŰ�</td>
      	<td>
      		<%= FormatNumber(ooffitem.FOneItem.Fshopitemprice,0) %>
      		<!--
      		/ <%= FormatNumber(ooffitem.FOneItem.Fshopsuplycash,0) %>
      		-->
      		&nbsp;&nbsp;
      		<!--
      	    <% if ooffitem.FOneItem.Fshopitemprice<>0 then %>
			<%= CLng((1- ooffitem.FOneItem.Fshopsuplycash/ooffitem.FOneItem.Fshopitemprice)*100) %> %
			<% end if %>

			&nbsp;&nbsp;
			-->
			<!-- ���ο���/�������뿩�� -->
			<% if (ooffitem.FOneItem.FShopItemOrgprice>ooffitem.FOneItem.Fshopitemprice) then %>
			    <font color=red>
			    <% if (ooffitem.FOneItem.FShopItemOrgprice<>0) then %>
			        <%= CLng((ooffitem.FOneItem.FShopItemOrgprice-ooffitem.FOneItem.Fshopitemprice)/ooffitem.FOneItem.FShopItemOrgprice*100) %> %
			    <% end if %>
			     ����
			    </font>
			<% end if %>


      	</td>
      	<td>��������</td>
      	<td colspan=2>

		</td>
    </tr>

    	<tr bgcolor="#FFFFFF">
	      	<td>�ɼǸ�</td>
	      	<td><%= ooffitem.FOneItem.Fshopitemoptionname %></td>
	      	<td></td>
	      	<td></td>
	      	<td></td>
	    </tr>

</table>
<% else %>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
        <td align="center">[�˻� ����� �����ϴ�.]</td>
    </tr>
</table>
<% end if %>
<p>
<% if osummarystock.FResultCount>0 then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name=calcufrm >
	<input type="hidden" name="orgrealstock" value="<%= osummarystock.FOneItem.Frealstock %>">
	<input type="hidden" name="orgerrrealcheckno" value="<%= osummarystock.FOneItem.Ferrrealcheckno %>">
	<input type="hidden" name="availsysstock" value="<%= osummarystock.FOneItem.Favailsysstock %>">
	<input type="hidden" name="todaybaljuno" value="<%= osummarystock.FOneItem.GetTodayBaljuNo %>">
	<input type="hidden" name="todayinputedrealcheckerrno" value="<%= otodayerritem.FOneItem.Ferrrealcheckno %>">

<!-- �ǽð� ������Ʈ ��
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td colspan="16" align=right>����������Ʈ : <%= osummarystock.FOneItem.Flastupdate %> </td>
    </tr>
-->
    <tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50">���԰�/��ǰ</td>
    	<td width="50">���Ǹ�/��ǰ</td>
		<td width="50">�����/��ǰ</td>
		<td width="50">��Ÿ���/��ǰ</td>
		<td width="50">CS<br>���/��ǰ</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">�ý���<br>�����</td>
		<td width="50">�ѽǻ�<br>����</td>
		<td width="50">�Ѻҷ�</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">�ǻ�<br>��ȿ���</td>
		<td width="50">ON��ǰ<br>�غ�</td>
		<td width="50">OFF��ǰ<br>�غ�</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">����ľ�<br>���</td>
		<td width="50">ON����<br>�Ϸ�</td>
		<td width="50">ON�ֹ�<br>����</td>
		<td width="50">OFF�ֹ�<br>����</td>
    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td rowspan="2"><%= osummarystock.FOneItem.Ftotipgono %></td>
    	<td rowspan="2"><%= -1*osummarystock.FOneItem.Ftotsellno %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Foffchulgono + osummarystock.FOneItem.Foffrechulgono %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Fetcchulgono + osummarystock.FOneItem.Fetcrechulgono %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Ferrcsno %></td>
    	<td rowspan="2" bgcolor="<%= adminColor("gray") %>"><%= osummarystock.FOneItem.Ftotsysstock %></td>
    	<td><%= osummarystock.FOneItem.Ferrrealcheckno %></td>
    	<td rowspan="2"><%= osummarystock.FOneItem.Ferrbaditemno %></td>
    	<td rowspan="2" bgcolor="<%= adminColor("gray") %>"><%= osummarystock.FOneItem.Frealstock %></td>
    	<td><%= osummarystock.FOneItem.Fipkumdiv5 %></td>
    	<td><%= osummarystock.FOneItem.Foffconfirmno %></td>
    	<td rowspan="2" bgcolor="<%= adminColor("gray") %>"><input type="text" name="checkstock" value="<%= osummarystock.FOneItem.GetCheckStockNo %>" size="4" maxlength="7" style="border:1px #999999 solid; text-align=center" onKeyUp="RecalcuErr();"></td>
    	<td><%= osummarystock.FOneItem.Fipkumdiv4 %></td>
    	<td><%= osummarystock.FOneItem.Fipkumdiv2 %></td>
    	<td><%= osummarystock.FOneItem.Foffjupno %></td>

    </tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td ><input type="text" name="errrealcheckno" value="<%= osummarystock.FOneItem.Ferrrealcheckno  %>"  size="4" maxlength="7" readonly style="background:#CCCCCC; border:1px #999999 solid; text-align=center"></td>
    	<td colspan="2"><%= osummarystock.FOneItem.Fipkumdiv5 + osummarystock.FOneItem.Foffconfirmno %></td>
    	<td colspan="3"><%= osummarystock.FOneItem.Fipkumdiv4 + osummarystock.FOneItem.Fipkumdiv2 + osummarystock.FOneItem.Foffjupno %></td>

    </tr>
    <tr bgcolor="#FFFFFF">
    	<td colspan="6" align=right>���� �Էµ� ����</td>
    	<td align="center" ><input type="text" name="todayerrrealcheckno" value="<%= otodayerritem.FOneItem.Ferrrealcheckno %>"  size="4" maxlength="7" readonly style="background:#CCCCCC; border:1px #999999 solid; text-align=center"></td>
		<td colspan="4"></td>
		<td align="left" colspan="4" >
		<input type="button" class="button" value="�ǻ��������" onclick="SaveErr();" <%= ChkIIF((Not IsVaildCode) And (Not IsStockExists),"disabled","") %> >
		</td>
	</tr>
	</form>
</table>

<form name=frmrefresh method=post action="/admin/stock/stockrefresh_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="realstock" value="">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="<%= itemoption %>">
</form>
<% end if %>
<%
set oitem = Nothing
set oitemoption = Nothing
set ooffitem = Nothing
set otodayerritem = Nothing
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->