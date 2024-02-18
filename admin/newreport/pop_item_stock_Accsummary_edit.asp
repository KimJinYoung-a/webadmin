<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : ���� ��� ���� �� ����
' Hieditor : 2014.05.26 ������ ����
'			 2022.10.11 �ѿ�� ����(��������, ǥ���ڵ����� ����)
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/offitemstock_cls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->

<%
dim itemgubun, itemid, itemoption
dim yyyy1, mm1
dim i


itemgubun = requestCheckVar(trim(request("itemgubun")),32)
itemid = requestCheckVar(getNumeric(trim(request("itemid"))),10)
itemoption = requestCheckVar(trim(request("itemoption")),32)

yyyy1 = requestCheckVar(getNumeric(trim(request("yyyy1"))),32)
mm1 = requestCheckVar(getNumeric(trim(request("mm1"))),32)

dim oitem
if itemgubun = "10" then
	set oitem = new CItemInfo
	oitem.FRectItemID = itemid

	if itemid<>"" then
		oitem.GetOneItemInfo
	end if

else
	set oitem = new CoffstockItemlist	'//�¶��� ��ũ������� Ŭ������ �浹, �������� ���� ����
	oitem.frectitemgubun = itemgubun
	oitem.FRectItemID = itemid
	oitem.frectitemoption = itemoption

	if itemid<>"" then
		oitem.GetoffItemDefaultData
	end if
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid

if itemid<>"" and itemgubun="10" then
	oitemoption.GetItemOptionInfo
end if

if (oitemoption.FResultCount<1) then
    ''if (Not isShopreturnItem) then
    	itemoption = "0000"
    ''end if
end if

dim BasicMonth
BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)

dim osummaryMonthstock
set osummaryMonthstock = new CSummaryItemStock
	osummaryMonthstock.FRectYYYYMM = BasicMonth
	osummaryMonthstock.FRectItemGubun = itemgubun
	osummaryMonthstock.FRectItemID =  itemid
	osummaryMonthstock.FRectItemOption =  itemoption

	if itemid<>"" then
		osummaryMonthstock.GetMonthly_Logisstock_Summary
	end if

dim osummarystock, isCurrStockExists
set osummarystock = new CSummaryItemStock
	osummarystock.FRectStartDate = BasicMonth + "-01"
	osummarystock.FRectItemGubun = itemgubun
	osummarystock.FRectItemID =  itemid
	osummarystock.FRectItemOption =  itemoption

	if itemid<>"" then
		osummarystock.GetCurrentItemStock
		isCurrStockExists= (osummarystock.FResultCount>0)
		osummarystock.GetDaily_Logisstock_Summary
	end if

dim oLastMonthstock
set oLastMonthstock = new CSummaryItemStock
	oLastMonthstock.FRectItemGubun = itemgubun
	oLastMonthstock.FRectItemID =  itemid
	oLastMonthstock.FRectItemOption =  itemoption

	if itemid<>"" then
	   oLastMonthstock.getLastMonthStock
	end if

dim sum_ipgono,sum_reipgono,sum_sellno,sum_resellno
dim sum_offchulgono, sum_offrechulgono, sum_etcchulgono, sum_etcrechulgono
dim sum_totsysstock, sum_availsysstock, sum_realstock
dim sum_errbaditemno, sum_errrealcheckno, sum_errcsno
dim mm_ipgono,mm_reipgono,mm_sellno,mm_resellno ,sysstock, sysavailstock, realstock, maystock ,ErrMsg, realstockWithBad
dim mm_offchulgono, mm_offrechulgono, mm_etcchulgono, mm_etcrechulgono ,mm_errbaditemno, mm_errrealcheckno, mm_errcsno


dim oCMonthlyStockLogics
set oCMonthlyStockLogics = new CMonthlyStock
	oCMonthlyStockLogics.FRectItemGubun = itemgubun
	oCMonthlyStockLogics.FRectItemID =  itemid
	oCMonthlyStockLogics.FRectItemOption =  itemoption

	if itemid<>"" then
	   oCMonthlyStockLogics.GetMonthlyMWDivHistoryLogics
	end if

dim oCMonthlyStockShop
set oCMonthlyStockShop = new CMonthlyStock
	oCMonthlyStockShop.FRectItemGubun = itemgubun
	oCMonthlyStockShop.FRectItemID =  itemid
	oCMonthlyStockShop.FRectItemOption =  itemoption

	if itemid<>"" then
	   oCMonthlyStockShop.GetMonthlyMWDivHistoryShop
	end if

%>
<script type='text/javascript'>

function popAssignMonthlyAccMwgubun(yyyymm, stockPlace, shopid, itemgubun, itemid, itemoption) {
    alert('������ ����');
    <% if not (C_ADMIN_AUTH or C_OFF_AUTH or C_MngPart) then %>
        return;
    <% end if %>
    var iURL = "/admin/newreport/popAssignMonthlyAccMwgubun.asp?yyyymm=" + yyyymm + "&stockPlace=" + stockPlace + "&shopid=" + shopid + "&itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption;
    var popwin = window.open(iURL,'popAssignMonthlyAccMwgubun','scrollbas=yes,resizable=yes,width=500,height=400');
    popwin.focus();
}

function popAssignMonthlyAccCenterMwgubun(yyyymm, stockPlace, shopid, itemgubun, itemid, itemoption) {
    alert('������ ����');
    <% if not (C_ADMIN_AUTH or C_OFF_AUTH or C_MngPart) then %>
        return;
    <% end if %>
    var iURL = "/admin/newreport/popAssignMonthlyAccCenterMwgubun.asp?yyyymm=" + yyyymm + "&stockPlace=" + stockPlace + "&shopid=" + shopid + "&itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption;
    var popwin = window.open(iURL,'popAssignMonthlyAccCenterMwgubun','scrollbas=yes,resizable=yes,width=500,height=400');
    popwin.focus();
}

function popAssignMonthlyAccVAT(yyyymm, stockPlace, shopid, itemgubun, itemid, itemoption) {
    alert('������ ����');
    <% if not (C_ADMIN_AUTH or C_OFF_AUTH or C_MngPart) then %>
        return;
    <% end if %>
    var iURL = "/admin/newreport/popAssignMonthlyAccVAT.asp?yyyymm=" + yyyymm + "&stockPlace=" + stockPlace + "&shopid=" + shopid + "&itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption;
    var popwin = window.open(iURL,'popAssignMonthlyAccVAT','scrollbas=yes,resizable=yes,width=500,height=400');
    popwin.focus();
}

function popAssignMonthlyAccLastIpgo(yyyymm, stockPlace, shopid, itemgubun, itemid, itemoption) {
    alert('������ ����');
    <% if not (C_ADMIN_AUTH or C_OFF_AUTH or C_MngPart) then %>
        return;
    <% end if %>
    var iURL = "/admin/newreport/popAssignMonthlyAccLastIpgo.asp?yyyymm=" + yyyymm + "&stockPlace=" + stockPlace + "&shopid=" + shopid + "&itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption;
    var popwin = window.open(iURL,'popAssignMonthlyAccLastIpgo','scrollbas=yes,resizable=yes,width=500,height=400');
    popwin.focus();
}

function popAssignMonthlyAccPrice(yyyymm, stockPlace, shopid, itemgubun, itemid, itemoption) {
    alert('������ ����');
    <% if not (C_ADMIN_AUTH or C_OFF_AUTH or C_MngPart) then %>
        return;
    <% end if %>
    var iURL = "/admin/newreport/popAssignMonthlyAccPrice.asp?yyyymm=" + yyyymm + "&stockPlace=" + stockPlace + "&shopid=" + shopid + "&itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption;
    var popwin = window.open(iURL,'popAssignMonthlyAccPrice','scrollbas=yes,resizable=yes,width=500,height=400');
    popwin.focus();
}

function popAssignMonthlyAccMakerid(yyyymm, stockPlace, shopid, itemgubun, itemid, itemoption) {
    alert('������ ����');
    <% if not (C_ADMIN_AUTH or C_OFF_AUTH or C_MngPart) then %>
        return;
    <% end if %>
    var iURL = "/admin/newreport/popAssignMonthlyAccMakerid.asp?yyyymm=" + yyyymm + "&stockPlace=" + stockPlace + "&shopid=" + shopid + "&itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption;
    var popwin = window.open(iURL,'popAssignMonthlyAccMakerid','scrollbas=yes,resizable=yes,width=500,height=400');
    popwin.focus();
}

function jsSaveExcStock(itemgubun, itemid, itemoption) {
    alert('������ ����');
    <% if not (C_ADMIN_AUTH or C_OFF_AUTH or C_MngPart) then %>
        return;
    <% end if %>
    var iURL = "/admin/newreport/popAssignMonthlyAccExc.asp?itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption;
    var popwin = window.open(iURL,'jsSaveExcStock','scrollbas=yes,resizable=yes,width=500,height=400');
    popwin.focus();
}

function refreshAccStock(comp,yyyymm,itemgubun, itemid, itemoption){
	var frm =document.frmRefresh;
	frm.mode.value = "itemAccStock";
	frm.yyyymm.value = yyyymm;
	frm.itemgubun.value = itemgubun;
	frm.itemid.value = itemid;
	frm.itemoption.value = itemoption;

	if (confirm(yyyymm+'�� �⸻��� ���ΰ�ħ �Ͻðڽ��ϱ�?')){
		comp.disabled=true;
		frm.submit();
	}
}

function refreshAccStockShop(comp,yyyymm,shopid,itemgubun, itemid, itemoption){
	var frm =document.frmRefresh;
	frm.mode.value = "itemAccStockShop";
	frm.yyyymm.value = yyyymm;
	frm.shopid.value = shopid;
	frm.itemgubun.value = itemgubun;
	frm.itemid.value = itemid;
	frm.itemoption.value = itemoption;

	var confirmstr = yyyymm+'�� ������ü �⸻��� ���ΰ�ħ �Ͻðڽ��ϱ�?'
	if (shopid!="") confirmstr = yyyymm+'�� '+shopid+' ���� �⸻��� ���ΰ�ħ �Ͻðڽ��ϱ�?'

	if (confirm(confirmstr)){
		comp.disabled=true;
		frm.submit();
	}
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="#FFFFFF" >
			<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�</td>
			<td align="left">
				<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
					<tr align="center" bgcolor="#FFFFFF" >
						<td align="left">
					��ǰ�ڵ�:
		        			<select class="select" name="itemgubun">
		        				<option value="10" <%= chkIIF(itemgubun="10","selected","") %> >10</option>
		        				<option value="70" <%= chkIIF(itemgubun="70","selected","") %> >70</option>
								<option value="75" <%= chkIIF(itemgubun="75","selected","") %> >75</option>
		        				<option value="80" <%= chkIIF(itemgubun="80","selected","") %> >80</option>
								<option value="85" <%= chkIIF(itemgubun="85","selected","") %> >85</option>
		        				<option value="90" <%= chkIIF(itemgubun="90","selected","") %> >90</option>
								<option value="98" <%= chkIIF(itemgubun="98","selected","") %> >98</option>
		        			</select>

		        			<input type="text" class="text" name=itemid value="<%= itemid %>" size=8 maxlength=8  onKeyPress="if (event.keyCode == 13){ document.frm.submit(); return false;}">

		        			<input type="text" class="text_ro" value="<%= itemoption %>" size=4 maxlength=4 readonly>

							<% if oitemoption.FResultCount>0 then %>

							<select class="select" name="itemoption">
								<option value="0000">----
									<% for i=0 to oitemoption.FResultCount-1 %>
									<option value="<%= oitemoption.FITemList(i).FItemOption %>" <% if itemoption=oitemoption.FITemList(i).FItemOption then response.write "selected" %> ><%= oitemoption.FITemList(i).FOptionName %>
										<% next %>
							</select>
							<% end if %>

		        			<input type="button" class="button" value="�˻�" onclick="document.frm.submit();">
						</td>
						<td align="right">
							<% if (FALSE) then %>
								<% if oitem.FResultCount>0 or (isCurrStockExists) then %>
									<% if itemid<>"" then %>
									����������Ʈ : <b><%= osummarystock.FOneItem.Flastupdate %></b>
									<% end if %>
									<% if (C_ADMIN_AUTH=true) or (session("ssBctId")="josin222") or (session("ssBctId")="fantasiax") then %>

										<% if (session("ssBctId")<>"fantasiax") then %>
										<input type="button" class="button" value="����� ��ü ���ΰ�ħ" onclick="RefreshIpchulStock();">
										<% end if %>

										<% if session("ssBctId")="icommang" then %>
										<!-- <input type="button" class="button" value="�Ǹų�����ü���ΰ�ħ" onclick="RefreshOldTotalSellStock();"> -->
										<% end if %>
									<% end if %>
									<input type="button" class="button" value="���ΰ�ħ" onclick="RefreshRecentStock();">
								<% end if %>
							<% end if %>

							<% if (C_ADMIN_AUTH or C_OFF_AUTH) then %>
							<input type="button" class="button" value="����ڻ꿡�� ����" onclick="jsSaveExcStock('<%= itemgubun %>', '<%= itemid %>', '<%= itemoption %>');" style="width:120px">
							<% end if %>
		    			</td>
					</tr>
				</table>
			</td>
		</tr>
</table>
</form>

<br>

<% if (oitem.FResultCount>0) then %>
<% if itemgubun="10" then %>
	<% if (oitem.FResultCount>0) then %>
	<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td rowspan=<%= 6 + oitemoption.FResultCount -1 %> width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
		  	<td width="60">��ǰ�ڵ�</td>
		  	<td width="300">
		  		<%= itemgubun %> <b><%= CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) %></b> <%= itemoption %>
		  		&nbsp;
		  		<% if (FALSE) and itemgubun="10" then %>
		  		<input type="button" class="button" value="����" onclick="PopItemSellEdit('<%= itemid %>');">
		  		<% end if %>
		  	</td>
		  	<td colspan="5">��չ�ۼҿ��� :
				<% if (oitem.FOneItem.FavgDLvDate>-1) then %>
			    <a href="javascript:popItemAvgDlvList('<%= itemid %>');">D+<%= oitem.FOneItem.FavgDLvDate+1 %></a>
				<% else %>
			    <a href="javascript:popItemAvgDlvList('<%= itemid %>');">������ ����</a>
				<% end if %>
			</td>

		</tr>
		<tr bgcolor="#FFFFFF">
		  	<td>�귣��ID</td>
		  	<td><%= oitem.FOneItem.FMakerid %></td>
		  	<td>�Ǹſ���</td>
		  	<td colspan=4><font color="<%= ynColor(oitem.FOneItem.FSellyn) %>"><%= oitem.FOneItem.FSellyn %></font></td>
		</tr>
		<tr bgcolor="#FFFFFF">
		  	<td>��ǰ��</td>
		  	<td><%= oitem.FOneItem.FItemName %></td>
		  	<td>��뿩��</td>
		  	<td colspan=4><font color="<%= ynColor(oitem.FOneItem.FIsUsing) %>"><%= oitem.FOneItem.FIsUsing %></font></td>
		</tr>
		<tr bgcolor="#FFFFFF">
		  	<td>�ǸŰ�</td>
		  	<td>
		  		<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
		  		&nbsp;&nbsp;
		  		<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
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
		  	<td colspan="2">
		  		<%= fncolor(oitem.FOneItem.Fdanjongyn,"dj") %>
		  		<% if oitem.FOneItem.Fdanjongyn="N" then %>
				������
				<% end if %>
			</td>
			<td align="center">���ڵ�</td>
			<td align="center">��ü�ڵ�</td>
		</tr>

		<% if oitemoption.FResultCount>1 then %>
		<!-- �ɼ��� �ִ°�� -->
		<% for i=0 to oitemoption.FResultCount -1 %>
		<% if oitemoption.FITemList(i).FOptIsUsing<>"Y" then %>
		<tr bgcolor="#FFFFFF">
			<td><font color="#AAAAAA">�ɼǸ� :</font></td>
			<td><font color="#AAAAAA"><%
			Response.Write "[" & oitemoption.FITemList(i).Fitemoption & "]" & oitemoption.FITemList(i).FOptionName & "&nbsp;"
			Response.Write CHKIIF(oitemoption.FITemList(i).Foptaddprice <> "0","(+"&FormatNumber(oitemoption.FITemList(i).Foptaddprice,0)&")","")
			      					  %></font></td>
			<td><font color="#AAAAAA">�������� : </font></td>
			<td><font color="#AAAAAA"><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</font></td>
			<td>���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
			<td align="center"><%= oitemoption.FITemList(i).Fbarcode %></td>
			<td align="center"><%= oitemoption.FITemList(i).Fupchemanagecode %></td>
		</tr>
		<% else %>

		<% if oitemoption.FITemList(i).Fitemoption=itemoption then %>
		<tr bgcolor="#EEEEEE">
			<% else %>
			<tr bgcolor="#FFFFFF">
			    <% end if %>
			    <td>�ɼǸ�</td>
			    <td><%
			    Response.Write "[" & oitemoption.FITemList(i).Fitemoption & "]" & oitemoption.FITemList(i).FOptionName & "&nbsp;"
			    Response.Write CHKIIF(oitemoption.FITemList(i).Foptaddprice <> "0","(+"&FormatNumber(oitemoption.FITemList(i).Foptaddprice,0)&")","")
			      	%></td>
			    <td>��������</td>
			    <td><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
			    <td>
			      	  ���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)
				    <% if (oitem.FOneItem.Fdanjongyn = "S") then %>
				      (���԰����� : <%= oitemoption.FITemList(i).Frestockdate %>)
				    <% end if %>
			    </td>
				<td align="center"><%= oitemoption.FITemList(i).Fbarcode %></td>
				<td align="center"><%= oitemoption.FITemList(i).Fupchemanagecode %></td>
			</tr>
			<% end if %>
		    <% next %>
			<% else %>
			<tr bgcolor="#FFFFFF">
		      	<td>�ɼǸ�</td>
		      	<td>-</td>
		      	<td>��������</td>
		      	<td><font color="<%= ynColor(oitem.FOneItem.Flimityn) %>"><%= oitem.FOneItem.Flimityn %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
		      	<td>
		      		���� ����� (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)
					<% if ((oitem.FOneItem.Fdanjongyn="S") and (oitemoption.FResultCount<1)) then %>

					<% end if %>
		      	</td>
				<td align="center"><%= oitem.FOneItem.Fbarcode %></td>
				<td align="center"><%= oitem.FOneItem.Fupchemanagecode %></td>
		    </tr>
			<% end if %>
	</table>
	<% end if %>
<%
'//�¶��� ���� ������
else
%>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#FFFFFF">
			<td rowspan=<%= 5 + oitem.FResultCount -1 %> width="110" valign="top" align="center">
				<img src="<%= oitem.foneitem.FImageList %>" width="100" height="100">
			</td>
	  		<td width="60"><b>*��ǰ����</b></td>
	  		<td width="300">
	  			<!--<input type="button" value="����" onclick="pop_itemedit_off_edit('<%'= oitem.foneitem.Fitemgubun %><%'=  Format00(6,oitem.foneitem.Fitemid) %><%'= oitem.foneitem.Fitemoption %>');" class="button">-->
	  		</td>
	  		<td width="60">�귣��ID :</td>
	  		<td colspan=2><%= oitem.foneitem.FMakerid %></td>
		</tr>
		<tr bgcolor="#FFFFFF">
	  		<td>��ǰ�ڵ� :</td>
	  		<td><%= oitem.foneitem.fitemgubun %> <b><%= CHKIIF(oitem.foneitem.FItemID>=1000000,Format00(8,oitem.foneitem.FItemID),Format00(6,oitem.foneitem.FItemID)) %></b> <%= oitem.foneitem.fitemoption %></td>
	  		<td>��뿩�� : </td>
	  		<td colspan=2><%= oitem.foneitem.FIsUsing %></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td>�ǸŰ� :</td>
			<td>
				<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
			</td>
	  		<td>��ǰ�� :</td>
	  		<td><%= oitem.foneitem.FItemName %></td>
		</tr>
		<tr bgcolor="#FFFFFF">
      		<td><font color="#AAAAAA">�ɼǸ� :</font></td>
      		<td><font color="#AAAAAA"><%= oitem.foneitem.FItemOptionName %></font></td>
      		<td><font color="#AAAAAA">������� : </font></td>
      		<td>
      			<%= oitem.foneitem.GetCheckStockNo %> : (NEW)
      		</td>
		</tr>
	</table>
<% end if %>
<% end if %>

<br>

<% if (oitem.FResultCount>0) or (itemgubun<>"10" and osummaryMonthstock.FResultCount>0)  then %>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<b>*�Ϻ� ���⳻��</b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="80">�Ͻ�</td>
      	<td width="55">�԰�</td>
      	<td width="55">��ǰ</td>
      	<td width="55">ON<br>���</td>
      	<td width="55">ON<br>��ǰ</td>
      	<td width="55">OFF<br>���</td>
      	<td width="55">OFF<br>��ǰ</td>

      	<td width="55">��Ÿ<br>���/��ǰ</td>
      	<td width="55">CS<br>���/��ǰ</td>
        <td width="60">�ý���<br>�����</td>
        <td width="55">(�ǻ�)<br>����</td>
        <td width="60">�ǻ�<br>���</td>
      	<td width="55">�ҷ�</td>
      	<!-- td width="60">�ý���<br>��ȿ���</td -->
      	<td width="60">�ǻ�<br>��ȿ���</td>
      	<td>���</td>
    </tr>
<!-- �����α� -->
<% if osummaryMonthstock.FResultCount>0 then %>
<% for i=0 to osummaryMonthstock.FResultCount-1 %>
<%
sum_ipgono = sum_ipgono + osummaryMonthstock.FItemList(i).Fipgono
sum_reipgono = sum_reipgono + osummaryMonthstock.FItemList(i).Freipgono
sum_sellno = sum_sellno + osummaryMonthstock.FItemList(i).Fsellno
sum_resellno = sum_resellno + osummaryMonthstock.FItemList(i).Fresellno
sum_offchulgono = sum_offchulgono + osummaryMonthstock.FItemList(i).Foffchulgono
sum_offrechulgono = sum_offrechulgono + osummaryMonthstock.FItemList(i).Foffrechulgono
sum_etcchulgono = sum_etcchulgono + osummaryMonthstock.FItemList(i).Fetcchulgono
sum_etcrechulgono = sum_etcrechulgono + osummaryMonthstock.FItemList(i).Fetcrechulgono
sum_errbaditemno	= sum_errbaditemno + osummaryMonthstock.FItemList(i).Ferrbaditemno
sum_errrealcheckno	= sum_errrealcheckno + osummaryMonthstock.FItemList(i).Ferrrealcheckno
sum_errcsno         = sum_errcsno + osummaryMonthstock.FItemList(i).Ferrcsno

sum_totsysstock = sum_totsysstock + osummaryMonthstock.FItemList(i).Ftotsysstock
sum_availsysstock = sum_availsysstock + osummaryMonthstock.FItemList(i).Favailsysstock
sum_realstock = sum_realstock + osummaryMonthstock.FItemList(i).Frealstock


sysstock = sysstock + osummaryMonthstock.FItemList(i).Ftotsysstock
sysavailstock = sysavailstock + osummaryMonthstock.FItemList(i).Favailsysstock
realstock = realstock + osummaryMonthstock.FItemList(i).Frealstock
maystock = maystock + osummaryMonthstock.FItemList(i).Frealstock

realstockWithBad = sysstock+sum_errrealcheckno ''2013/11/22�߰�

'sum_offsell = sum_offsell + osummaryMonthstock.FItemList(i).Foffsellno
'offstockno = offstockno + osummaryMonthstock.FItemList(i).Foffchulgono*-1 + osummaryMonthstock.FItemList(i).Foffrechulgono*-1 - osummaryMonthstock.FItemList(i).Foffsellno
%>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= osummaryMonthstock.FItemList(i).Fyyyymm %></td>
      	<td><a href="javascript:PopItemIpChulList('<%= osummaryMonthstock.FItemList(i).Fyyyymm %>-01','<%= DateSerial(Left(osummaryMonthstock.FItemList(i).Fyyyymm,4),Right(osummaryMonthstock.FItemList(i).Fyyyymm,2)+1,0) %>','<%= osummaryMonthstock.FItemList(i).Fitemgubun %>','<%= osummaryMonthstock.FItemList(i).Fitemid %>','<%= osummaryMonthstock.FItemList(i).FItemoption %>','I');"><%= osummaryMonthstock.FItemList(i).Fipgono %></a></td>
      	<td><%= osummaryMonthstock.FItemList(i).Freipgono %></td>
      	<td><a href="javascript:popBuyItemListChulgo('<%= osummaryMonthstock.FItemList(i).Fyyyymm %>')"><%= osummaryMonthstock.FItemList(i).Fsellno %></a></td>
      	<td><%= osummaryMonthstock.FItemList(i).Fresellno %></td>
      	<td><a href="javascript:PopItemIpChulList('<%= osummaryMonthstock.FItemList(i).Fyyyymm %>-01','<%= DateSerial(Left(osummaryMonthstock.FItemList(i).Fyyyymm,4),Right(osummaryMonthstock.FItemList(i).Fyyyymm,2)+1,0) %>','<%= osummaryMonthstock.FItemList(i).Fitemgubun %>','<%= osummaryMonthstock.FItemList(i).Fitemid %>','<%= osummaryMonthstock.FItemList(i).FItemoption %>','S');"><%= osummaryMonthstock.FItemList(i).Foffchulgono %></a></td>
      	<td><%= osummaryMonthstock.FItemList(i).Foffrechulgono %></td>

      	<td><a href="javascript:PopItemIpChulList('<%= osummaryMonthstock.FItemList(i).Fyyyymm %>-01','<%= DateSerial(Left(osummaryMonthstock.FItemList(i).Fyyyymm,4),Right(osummaryMonthstock.FItemList(i).Fyyyymm,2)+1,0) %>','<%= osummaryMonthstock.FItemList(i).Fitemgubun %>','<%= osummaryMonthstock.FItemList(i).Fitemid %>','<%= osummaryMonthstock.FItemList(i).FItemoption %>','E');"><%= osummaryMonthstock.FItemList(i).Fetcchulgono + osummaryMonthstock.FItemList(i).Fetcrechulgono %></a></td>
    	<td><a href="javascript:popCSItemListChulgo('<%= osummaryMonthstock.FItemList(i).Fyyyymm %>')"><%= osummaryMonthstock.FItemList(i).Ferrcsno %></a></td>
        <td><%= sysstock %></td>
        <td><a href="javascript:popRealErrList('<%= osummaryMonthstock.FItemList(i).Fyyyymm %>-01','<%= DateSerial(Left(osummaryMonthstock.FItemList(i).Fyyyymm,4),Right(osummaryMonthstock.FItemList(i).Fyyyymm,2)+1,0) %>','<%= osummaryMonthstock.FItemList(i).Fitemgubun %>','<%= osummaryMonthstock.FItemList(i).Fitemid %>','<%= osummaryMonthstock.FItemList(i).FItemoption %>')"><%= osummaryMonthstock.FItemList(i).Ferrrealcheckno %></a></td>
      	<td><%= realstockWithBad %></td>
      	<td><a href="javascript:PopStockBaditem('<%= osummaryMonthstock.FItemList(i).Fyyyymm %>-01','<%= DateSerial(Left(osummaryMonthstock.FItemList(i).Fyyyymm,4),Right(osummaryMonthstock.FItemList(i).Fyyyymm,2)+1,0) %>','<%= osummaryMonthstock.FItemList(i).Fitemgubun %>','<%= osummaryMonthstock.FItemList(i).Fitemid %>','<%= osummaryMonthstock.FItemList(i).FItemoption %>')"><%= osummaryMonthstock.FItemList(i).Ferrbaditemno %></a></td>
      	<!-- td><%= sysavailstock %></td -->
      	<td><%= realstock %></td>
      	<td>
      	    <% if realstock<>0 then %>
      	    <%= CLng((osummaryMonthstock.FItemList(i).Fsellno + osummaryMonthstock.FItemList(i).Foffchulgono)*-1/realstock*100)/100 %>
      	    <% end if %>
      	</td>
    </tr>
<% next %>
	<tr align="center" bgcolor="#EEEEEE">
		<td>�����Ұ�</td>
		<td>
		    <%= sum_ipgono %>
		    <% if oLastMonthstock.FOneItem.Fipgono<>sum_ipgono then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Fipgono %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_reipgono %>
		    <% if oLastMonthstock.FOneItem.Freipgono<>sum_reipgono then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Freipgono %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_sellno %>
		    <% if oLastMonthstock.FOneItem.Fsellno<>sum_sellno then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Fsellno %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_resellno %>
		    <% if oLastMonthstock.FOneItem.Fresellno<>sum_resellno then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Fresellno %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_offchulgono %>
		    <% if oLastMonthstock.FOneItem.Foffchulgono<>sum_offchulgono then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Foffchulgono %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_offrechulgono %>
		    <% if oLastMonthstock.FOneItem.Foffrechulgono<>sum_offrechulgono then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Foffrechulgono %>)</font>
		    <% end if %>
		</td>

		<td>
		    <%= sum_etcchulgono + sum_etcrechulgono %>
		    <% if (oLastMonthstock.FOneItem.Fetcchulgono+oLastMonthstock.FOneItem.Fetcrechulgono)<>(sum_etcchulgono + sum_etcrechulgono) then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Fetcchulgono+oLastMonthstock.FOneItem.Fetcrechulgono %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_errcsno %>
		    <% if oLastMonthstock.FOneItem.Ferrcsno<>sum_errcsno then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Ferrcsno %>)</font>
		    <% end if %>
		</td>
		<td>
		    <b><%= sum_totsysstock %></b>
		    <% if oLastMonthstock.FOneItem.Ftotsysstock<>sum_totsysstock then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Ftotsysstock %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_errrealcheckno %>
		    <% if oLastMonthstock.FOneItem.Ferrrealcheckno<>sum_errrealcheckno then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Ferrrealcheckno %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_totsysstock+sum_errrealcheckno %>
		    <% if oLastMonthstock.FOneItem.Ftotsysstock+oLastMonthstock.FOneItem.Ferrrealcheckno<>sum_totsysstock+sum_errrealcheckno then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Ftotsysstock+oLastMonthstock.FOneItem.Ferrrealcheckno %>)</font>
		    <% end if %>
		</td>
		<td>
		    <%= sum_errbaditemno %>
		    <% if oLastMonthstock.FOneItem.Ferrbaditemno<>sum_errbaditemno then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Ferrbaditemno %>)</font>
		    <% end if %>
		</td>
		<!--
		<td>
		    <b><%= sum_availsysstock %></b>
		    <% if oLastMonthstock.FOneItem.Favailsysstock<>sum_availsysstock then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Favailsysstock %>)</font>
		    <% end if %>
		</td>
		-->
		<td>
		    <b><%= sum_realstock %></b>
		    <% if oLastMonthstock.FOneItem.Frealstock<>sum_realstock then %>
		    <br><font color="red">(<%= oLastMonthstock.FOneItem.Frealstock %>)</font>
		    <% end if %>
		</td>
		<td>
		</td>
	</tr>
<% end if %>
<!-- �Ϻ� �α� -->
<%
dim ismidSubtotalShow
%>
<% for i=0 to osummarystock.FResultCount-1 %>
<%
sum_ipgono = sum_ipgono + osummarystock.FItemList(i).Fipgono
sum_reipgono = sum_reipgono + osummarystock.FItemList(i).Freipgono
sum_sellno = sum_sellno + osummarystock.FItemList(i).Fsellno
sum_resellno = sum_resellno + osummarystock.FItemList(i).Fresellno
sum_offchulgono = sum_offchulgono + osummarystock.FItemList(i).Foffchulgono
sum_offrechulgono = sum_offrechulgono + osummarystock.FItemList(i).Foffrechulgono
sum_etcchulgono = sum_etcchulgono + osummarystock.FItemList(i).Fetcchulgono
sum_etcrechulgono = sum_etcrechulgono + osummarystock.FItemList(i).Fetcrechulgono
sum_errbaditemno	= sum_errbaditemno + osummarystock.FItemList(i).Ferrbaditemno
sum_errrealcheckno	= sum_errrealcheckno + osummarystock.FItemList(i).Ferrrealcheckno
sum_errcsno = sum_errcsno + osummarystock.FItemList(i).Ferrcsno
sum_totsysstock = sum_totsysstock + osummarystock.FItemList(i).Ftotsysstock
sum_availsysstock = sum_availsysstock + osummarystock.FItemList(i).Favailsysstock
sum_realstock = sum_realstock + osummarystock.FItemList(i).Frealstock

sysstock = sysstock + osummarystock.FItemList(i).Ftotsysstock
sysavailstock = sysavailstock + osummarystock.FItemList(i).Favailsysstock
realstock = realstock + osummarystock.FItemList(i).Frealstock
maystock = maystock + osummarystock.FItemList(i).Frealstock


mm_ipgono = mm_ipgono + osummarystock.FItemList(i).Fipgono
mm_reipgono = mm_reipgono + osummarystock.FItemList(i).Freipgono
mm_sellno = mm_sellno + osummarystock.FItemList(i).Fsellno
mm_resellno = mm_resellno + osummarystock.FItemList(i).Fresellno
mm_offchulgono = mm_offchulgono + osummarystock.FItemList(i).Foffchulgono
mm_offrechulgono = mm_offrechulgono + osummarystock.FItemList(i).Foffrechulgono
mm_etcchulgono = mm_etcchulgono + osummarystock.FItemList(i).Fetcchulgono
mm_etcrechulgono = mm_etcrechulgono + osummarystock.FItemList(i).Fetcrechulgono
mm_errbaditemno	= mm_errbaditemno + osummarystock.FItemList(i).Ferrbaditemno
mm_errrealcheckno	= mm_errrealcheckno + osummarystock.FItemList(i).Ferrrealcheckno
mm_errcsno  = mm_errcsno + osummarystock.FItemList(i).Ferrcsno

'sum_offsell = sum_offsell + osummarystock.FItemList(i).Foffsellno
'offstockno = offstockno + osummarystock.FItemList(i).Foffchulgono*-1 + osummarystock.FItemList(i).Foffrechulgono*-1 - osummarystock.FItemList(i).Foffsellno
%>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= osummarystock.FItemList(i).Fyyyymmdd %>(<%= osummarystock.FItemList(i).GetDpartName %>)</td>
      	<td><a href="javascript:PopItemIpChulList('<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fitemgubun %>','<%= osummarystock.FItemList(i).Fitemid %>','<%= osummarystock.FItemList(i).FItemoption %>','I');"><%= osummarystock.FItemList(i).Fipgono %></a></td>
      	<td><%= osummarystock.FItemList(i).Freipgono %></td>
      	<td><a href="javascript:popBuyItemListChulgo('<%= osummarystock.FItemList(i).Fyyyymmdd %>');"><%= osummarystock.FItemList(i).Fsellno %></a></td>
      	<td><%= osummarystock.FItemList(i).Fresellno %></td>
      	<td><a href="javascript:PopItemIpChulList('<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fitemgubun %>','<%= osummarystock.FItemList(i).Fitemid %>','<%= osummarystock.FItemList(i).FItemoption %>','S');"><%= osummarystock.FItemList(i).Foffchulgono %></a></td>
      	<td><%= osummarystock.FItemList(i).Foffrechulgono %></td>

      	<td><a href="javascript:PopItemIpChulList('<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fitemgubun %>','<%= osummarystock.FItemList(i).Fitemid %>','<%= osummarystock.FItemList(i).FItemoption %>','E');"><%= osummarystock.FItemList(i).Fetcchulgono + osummarystock.FItemList(i).Fetcrechulgono %></a></td>
    	<td><a href="javascript:popCSItemListChulgo('<%= osummarystock.FItemList(i).Fyyyymmdd %>')"><%= osummarystock.FItemList(i).Ferrcsno %></a></td>
        <td><%= sysstock %></td>
        <td><a href="javascript:popRealErrList('<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fitemgubun %>','<%= osummarystock.FItemList(i).Fitemid %>','<%= osummarystock.FItemList(i).FItemoption %>')"><%= osummarystock.FItemList(i).Ferrrealcheckno %></a></td>
        <td><%= sysstock+sum_errrealcheckno %></td>
      	<td><a href="javascript:PopStockBaditem('<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fyyyymmdd %>','<%= osummarystock.FItemList(i).Fitemgubun %>','<%= osummarystock.FItemList(i).Fitemid %>','<%= osummarystock.FItemList(i).FItemoption %>')"><%= osummarystock.FItemList(i).Ferrbaditemno %></a></td>

      	<!-- td><%= sysavailstock %></td -->
      	<td><%= realstock %></td>
      	<td></td>

    </tr>
    <%
        ismidSubtotalShow = false

        if (i>=osummarystock.FResultCount-1) then
            ismidSubtotalShow = true
        elseif Left(osummarystock.FItemList(i).Fyyyymmdd,7)<>Left(osummarystock.FItemList(i+1).Fyyyymmdd,7) then
            ismidSubtotalShow = true
        end if

    %>
    <% if (ismidSubtotalShow) then %>
    <!-- ���� �հ� �߰� -->
    <tr align="center" bgcolor="#EEEEEE">
		<td><%= Left(osummarystock.FItemList(i).Fyyyymmdd,7) %></td>
		<td><%= mm_ipgono %></td>
		<td><%= mm_reipgono %></td>
		<td><%= mm_sellno %></td>
		<td><%= mm_resellno %></td>
		<td><%= mm_offchulgono %></td>
		<td><%= mm_offrechulgono %></td>

		<td><%= mm_etcchulgono + mm_etcrechulgono%></td>
		<td><%= mm_errcsno %></td>
        <td><b><%= sum_totsysstock %></b></td>
        <td><%= mm_errrealcheckno %></td>
        <td><%= sum_totsysstock+sum_errrealcheckno %></td>
		<td><%= mm_errbaditemno %></td>
		<!-- td><b><%= sum_availsysstock %></b></td -->
		<td><b><%= sum_realstock %></b></td>
        <td>
            <% if sum_realstock<>0 then %>
      	    <b><%= CLng((mm_sellno + mm_offchulgono)*-1/sum_realstock*100)/100 %></b>
      	    <% end if %>
        </td>
	</tr>
	<%
	mm_ipgono = 0
    mm_reipgono = 0
    mm_sellno = 0
    mm_resellno = 0
    mm_offchulgono = 0
    mm_offrechulgono = 0
    mm_etcchulgono = 0
    mm_etcrechulgono = 0
    mm_errbaditemno	= 0
    mm_errrealcheckno = 0
    mm_errcsno = 0
	%>
    <% end if %>
<% next %>
	<tr align="center" bgcolor="#EEEEEE">
		<td>ToTal</td>
		<td><%= sum_ipgono %></td>
		<td><%= sum_reipgono %></td>
		<td><%= sum_sellno %></td>
		<td><%= sum_resellno %></td>
		<td><%= sum_offchulgono %></td>
		<td><%= sum_offrechulgono %></td>

		<td><%= sum_etcchulgono + sum_etcrechulgono%></td>
		<td><%= sum_errcsno %></td>
        <td><b><%= sum_totsysstock %></b></td>
        <td><%= sum_errrealcheckno %></td>
        <td><%= sum_totsysstock+sum_errrealcheckno %></td>
		<td><%= sum_errbaditemno %></td>
		<!-- td><b><%= sum_availsysstock %></b></td -->
		<td><b><%= sum_realstock %></b></td>
        <td></td>

	</tr>
</table>

<br>

<% end if %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" align="center">
	<tr bgcolor="#FFFFFF" height="25">
		<td align="left" colspan="12"><b>*���� �������</b></td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center" height="25">
		<td width="60">����</td>
		<td width="120">����</td>
		<td width="80">���Ա���</td>
		<td width="80">��������</td>
		<td width="120">�귣��</td>
		<td width="80">���<br>���԰�</td>
		<td width="80">�ۼ���<br>���԰�</td>
		<td width="80">�����԰�</td>
		<td width="80">�����԰�</td>
		<td width="150">�����</td>
		<td width="150">��������</td>
		<td>���
		<% if (oCMonthlyStockLogics.FResultCount>0) then %>
			<% if (oCMonthlyStockLogics.FItemList(0).Fyyyymm<LEFT(now(),7)) then %>
				<% if (oCMonthlyStockLogics.FItemList(0).Fyyyymm<LEFT(dateadd("m",-1,now()),7)) then %>
				<input type="button" value="�⸻���ۼ� <%=LEFT(dateadd("m",-1,now()),7)%>" onClick="refreshAccStock(this,'<%=LEFT(dateadd("m",-1,now()),7)%>','<%=itemgubun%>','<%=itemid%>','<%=itemoption%>')">
				<% else %>
				<input type="button" value="�⸻���ۼ� <%=LEFT(dateadd("m",-0,now()),7)%>" onClick="refreshAccStock(this,'<%=LEFT(dateadd("m",-0,now()),7)%>','<%=itemgubun%>','<%=itemid%>','<%=itemoption%>')">
				<% end if %>
			<% end if %>
		<% end if %>
		</td>
	</tr>
	<% for i = 0 to oCMonthlyStockLogics.FResultCount - 1 %>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center"><%= oCMonthlyStockLogics.FItemList(i).Fyyyymm %></td>
		<td align="center"><%= oCMonthlyStockLogics.FItemList(i).Fshopid %></td>
		<td align="center">
			<a href="javascript:popAssignMonthlyAccMwgubun('<%= oCMonthlyStockLogics.FItemList(i).Fyyyymm %>', 'L', '<%= oCMonthlyStockLogics.FItemList(i).Fshopid %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemgubun %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemid %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemoption %>')">
				<%= oCMonthlyStockLogics.FItemList(i).getMaeipGubunName %>
			</a>
		</td>
		<td align="center">
			<a href="javascript:popAssignMonthlyAccVAT('<%= oCMonthlyStockLogics.FItemList(i).Fyyyymm %>', 'L', '<%= oCMonthlyStockLogics.FItemList(i).Fshopid %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemgubun %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemid %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemoption %>')">
				<%= oCMonthlyStockLogics.FItemList(i).Flastvatinclude %>
			</a>
		</td>
		<td align="center">
			<a href="javascript:popAssignMonthlyAccMakerid('<%= oCMonthlyStockLogics.FItemList(i).Fyyyymm %>', 'L', '<%= oCMonthlyStockLogics.FItemList(i).Fshopid %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemgubun %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemid %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemoption %>')">
				<%= oCMonthlyStockLogics.FItemList(i).Fmakerid %>
			</a>
		</td>
		<td align="right">
			<% if Not IsNull(oCMonthlyStockLogics.FItemList(i).FavgipgoPrice) then %>
			<a href="javascript:popAssignMonthlyAccPrice('<%= oCMonthlyStockLogics.FItemList(i).Fyyyymm %>', 'L', '<%= oCMonthlyStockLogics.FItemList(i).Fshopid %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemgubun %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemid %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemoption %>')">
			<%= FormatNumber(oCMonthlyStockLogics.FItemList(i).FavgipgoPrice, 0) %>
			</a>
			&nbsp;
			<% end if %>
		</td>
		<td align="right">
			<% if Not IsNull(oCMonthlyStockLogics.FItemList(i).FbuyPrice) then %>
			<a href="javascript:popAssignMonthlyAccPrice('<%= oCMonthlyStockLogics.FItemList(i).Fyyyymm %>', 'L', '<%= oCMonthlyStockLogics.FItemList(i).Fshopid %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemgubun %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemid %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemoption %>')">
				<%= FormatNumber(oCMonthlyStockLogics.FItemList(i).FbuyPrice, 0) %>
			</a>
			&nbsp;
			<% end if %>
		</td>
		<td align="center">
			<%= CHKIIF(IsNull(oCMonthlyStockLogics.FItemList(i).FfirstIpgoDate), "NULL", oCMonthlyStockLogics.FItemList(i).FfirstIpgoDate) %>
		</td>
		<td align="center">
			<a href="javascript:popAssignMonthlyAccLastIpgo('<%= oCMonthlyStockLogics.FItemList(i).Fyyyymm %>', 'L', '<%= oCMonthlyStockLogics.FItemList(i).Fshopid %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemgubun %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemid %>', '<%= oCMonthlyStockLogics.FItemList(i).Fitemoption %>')">
				<%= oCMonthlyStockLogics.FItemList(i).FlastIpgoDate %><%= CHKIIF(IsNull(oCMonthlyStockLogics.FItemList(i).FlastIpgoDate), "NULL", "") %>
			</a>
		</td>
		<td align="center"><%= oCMonthlyStockLogics.FItemList(i).Fregdate %></td>
		<td align="center">
			<% if (oCMonthlyStockLogics.FItemList(i).Fregdate <> oCMonthlyStockLogics.FItemList(i).Flastupdate) then %>
			<%= oCMonthlyStockLogics.FItemList(i).Flastupdate %>
			<% end if %>
		</td>
		<td align="center">
		<% if (oCMonthlyStockLogics.FItemList(i).Fyyyymm+"-01">=LEFT(dateadd("m",-1,LEFT(now(),7)+"-01"),10)) then %>
		<input type="button" value="�⸻���ۼ�" onClick="refreshAccStock(this,'<%=oCMonthlyStockLogics.FItemList(i).Fyyyymm%>','<%=itemgubun%>','<%=itemid%>','<%=itemoption%>')">
		<% end if %>
		</td>
	</tr>
	<% next %>
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" align="center">
	<tr bgcolor="#FFFFFF" height="25">
		<td align="left" colspan="12"><b>*���� �������</b></td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center" height="25">
		<td width="60">����</td>
		<td width="120">����</td>
		<td width="80">���Ա���</td>
		<td width="80">����<br />���Ա���</td>
		<td width="120">�귣��</td>
		<td width="80">���<br>���԰�</td>
		<td width="80">�ۼ���<br>���԰�</td>
		<td width="80">�����԰�<br />(�����԰�)</td>
		<td width="80">�����԰�<br />(�����԰�)</td>
		<td width="150">�����</td>
		<td width="150">��������</td>
		<td>���
		<% if (oCMonthlyStockShop.FResultCount>0) then %>
			<% if (oCMonthlyStockShop.FItemList(0).Fyyyymm<LEFT(now(),7)) then %>
				<% if (oCMonthlyStockShop.FItemList(0).Fyyyymm<LEFT(dateadd("m",-1,now()),7)) then %>
				<input type="button" value="�⸻���ۼ� <%=LEFT(dateadd("m",-1,now()),7)%>" onClick="refreshAccStockShop(this,'<%=LEFT(dateadd("m",-1,now()),7)%>','','<%=itemgubun%>','<%=itemid%>','<%=itemoption%>')">
				<% else %>
				<input type="button" value="�⸻���ۼ� <%=LEFT(dateadd("m",-0,now()),7)%>" onClick="refreshAccStockShop(this,'<%=LEFT(dateadd("m",-0,now()),7)%>','','<%=itemgubun%>','<%=itemid%>','<%=itemoption%>')">
				<% end if %>
			<% end if %>
		<% end if %>
		</td>
	</tr>
	<% for i = 0 to oCMonthlyStockShop.FResultCount - 1 %>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center"><%= oCMonthlyStockShop.FItemList(i).Fyyyymm %></td>
		<td align="center"><%= oCMonthlyStockShop.FItemList(i).Fshopid %></td>
		<td align="center">
			<a href="javascript:popAssignMonthlyAccMwgubun('<%= oCMonthlyStockShop.FItemList(i).Fyyyymm %>', 'S', '<%= oCMonthlyStockShop.FItemList(i).Fshopid %>', '<%= oCMonthlyStockShop.FItemList(i).Fitemgubun %>', '<%= oCMonthlyStockShop.FItemList(i).Fitemid %>', '<%= oCMonthlyStockShop.FItemList(i).Fitemoption %>')">
				<%= oCMonthlyStockShop.FItemList(i).Fmwdiv %>
				<% if IsNull(oCMonthlyStockShop.FItemList(i).Fmwdiv) then %>-<% end if %>
			</a>
		</td>
		<td align="center">
			<a href="javascript:popAssignMonthlyAccCenterMwgubun('<%= oCMonthlyStockShop.FItemList(i).Fyyyymm %>', 'S', '<%= oCMonthlyStockShop.FItemList(i).Fshopid %>', '<%= oCMonthlyStockShop.FItemList(i).Fitemgubun %>', '<%= oCMonthlyStockShop.FItemList(i).Fitemid %>', '<%= oCMonthlyStockShop.FItemList(i).Fitemoption %>')">
				<%= oCMonthlyStockShop.FItemList(i).Fcentermwdiv %>
				<%
				if IsNull(oCMonthlyStockShop.FItemList(i).Fcentermwdiv) then
					response.write "-"
				elseif Trim(oCMonthlyStockShop.FItemList(i).Fcentermwdiv) = "" then
					response.write "-"
				end if
				%>
			</a>
		</td>
		<td align="center">
			<a href="javascript:popAssignMonthlyAccMakerid('<%= oCMonthlyStockShop.FItemList(i).Fyyyymm %>', 'S', '<%= oCMonthlyStockShop.FItemList(i).Fshopid %>', '<%= oCMonthlyStockShop.FItemList(i).Fitemgubun %>', '<%= oCMonthlyStockShop.FItemList(i).Fitemid %>', '<%= oCMonthlyStockShop.FItemList(i).Fitemoption %>')">
				<%= oCMonthlyStockShop.FItemList(i).Fmakerid %>
			</a>
		</td>
		<td align="right">
			<% if Not IsNull(oCMonthlyStockShop.FItemList(i).FavgipgoPrice) then %>
			<a href="javascript:popAssignMonthlyAccPrice('<%= oCMonthlyStockShop.FItemList(i).Fyyyymm %>', 'S', '<%= oCMonthlyStockShop.FItemList(i).Fshopid %>', '<%= oCMonthlyStockShop.FItemList(i).Fitemgubun %>', '<%= oCMonthlyStockShop.FItemList(i).Fitemid %>', '<%= oCMonthlyStockShop.FItemList(i).Fitemoption %>')">
				<%= FormatNumber(oCMonthlyStockShop.FItemList(i).FavgipgoPrice, 0) %>
				&nbsp;
			</a>
			<% end if %>
		</td>
		<td align="right">
			<% if Not IsNull(oCMonthlyStockShop.FItemList(i).FbuyPrice) then %>
			<%= FormatNumber(oCMonthlyStockShop.FItemList(i).FbuyPrice, 0) %>
		&nbsp;
			<% end if %>
		</td>
		<td align="center"><%= oCMonthlyStockShop.FItemList(i).FlastIpgoDateLogics %></td>
		<td align="center"><%= oCMonthlyStockShop.FItemList(i).FlastIpgoDate %></td>
		<td align="center"><%= oCMonthlyStockShop.FItemList(i).Fregdate %></td>
		<td align="center">
			<% if (oCMonthlyStockShop.FItemList(i).Fregdate <> oCMonthlyStockShop.FItemList(i).Flastupdate) then %>
			<%= oCMonthlyStockShop.FItemList(i).Flastupdate %>
			<% end if %>
		</td>
		<td align="center">
		<% if (oCMonthlyStockShop.FItemList(i).Fyyyymm+"-01">=LEFT(dateadd("m",-1,LEFT(now(),7)+"-01"),10)) then %>
		<input type="button" value="�⸻���ۼ�" onClick="refreshAccStockShop(this,'<%=oCMonthlyStockShop.FItemList(i).Fyyyymm%>','<%=oCMonthlyStockShop.FItemList(i).Fshopid%>','<%=itemgubun%>','<%=itemid%>','<%=itemoption%>')">
		<% end if %>
		</td>
	</tr>
	<% next %>
</table>

<form name="frmRefresh" method="post" action="/admin/stock/stockrefresh_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="yyyymm" value="">
<input type="hidden" name="shopid" value="">
<input type="hidden" name="itemgubun" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
</form>

<%
set oCMonthlyStockShop=nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
