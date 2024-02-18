<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/realjaegocls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<%

const C_STOCK_DAY=7

dim itemgubun, itemid, itemoption
itemgubun = request("itemgubun")
itemid = request("itemid")
itemoption = request("itemoption")

if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"

dim ojaegoitem
set ojaegoitem = new CRealJaeGo
ojaegoitem.FRectItemID = itemid
if itemid<>"" then
	ojaegoitem.GetItemDefaultData
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



dim BasicMonth


BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)


dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectStartDate = BasicMonth + "-01"
osummarystock.FRectItemGubun = itemgubun
osummarystock.FRectItemID =  itemid
osummarystock.FRectItemOption =  itemoption
if itemid<>"" then
	osummarystock.GetCurrentItemStock
	osummarystock.GetDaily_Logisstock_Summary
end if

dim offstock
set offstock = new COffShopDailyStock
offstock.FRectItemGubun = itemgubun
offstock.FRectItemid = itemid
offstock.FRectItemoption = itemoption
if itemid<>"" then
	if ojaegoitem.FResultCount>0 then
		offstock.FRectMakerid = ojaegoitem.FItemList(0).Fmakerid
	end if

	offstock.GetCurrentAllShopItemStock
end if

dim i
dim sum_ipgono,sum_reipgono,sum_sellno,sum_resellno

dim sum_offchulgono, sum_offrechulgono, sum_etcchulgono, sum_etcrechulgono
dim sum_totsysstock, sum_availsysstock, sum_realstock
dim sum_errbaditemno, sum_errrealcheckno
dim sum_offsell

dim sysstock, sysavailstock, realstock, maystock
dim offstockno
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

function popRealErrInput(itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poprealerrinput.asp?itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&BasicMonth=<%= BasicMonth %>','poprealerrinput','width=900,height=460,scrollbar=yes,resizable=yes')
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
		<font color="red"><strong>OFFLINE��ǰ�������Ȳ</strong></font></td>
		<td height="20" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			�������� ���� �ǽð� ��ǰ��� �����Դϴ�..
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
        	��ǰ�ڵ�: <input type=text name=itemid value="<%= itemid %>" size=9 maxlength=9>
        	&nbsp;
			<% if oitemoption.FResultCount>0 then %>
			�ɼǼ��� :
			<select name="itemoption">
			<option value="0000">----
			<% for i=0 to oitemoption.FResultCount-1 %>
			<option value="<%= oitemoption.FItemList(i).FItemOption %>" <% if itemoption=oitemoption.FItemList(i).FItemOption then response.write "selected" %> ><%= oitemoption.FItemList(i).FItemOptionName %>
			<% next %>
			</select>
			<% end if %>
			&nbsp;
        	<input type=button value="�˻�" onclick="document.frm.submit();">
        </td>
        <td valign="top" align="right">
        <% if itemid<>"" then %>
        	����������Ʈ�ð� : <b><%= osummarystock.FOneItem.Flastupdate %></b>
        <% end if %>

        <% if C_ADMIN_AUTH=true then %>
        <!-- <input type="button" value="��ü�������ΰ�ħ" onclick="RefreshALLStock();"> -->
        <input type="button" value="2���� ���ΰ�ħ" onclick="RefreshRecentStock();" disabled >
        <% end if %>
        <input type="button" value="���� ���ΰ�ħ" onclick="RefreshTodayStock();" disabled >
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- ǥ ��ܹ� ��-->

<% if ojaegoitem.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#EEEEEE">
		<td colspan="6">&nbsp;<b> *Center ��� </b></td>
	</tr>
	<tr bgcolor="#FFFFFF">
    	<td rowspan=<%= 5 + ojaegoitem.FResultCount -1 %> width="110" valign=top align=center><img src="<%= ojaegoitem.FItemList(0).FImageList %>" width="100" height="100"></td>
      	<td width="60"><b>*��ǰ����</b></td>
      	<td width="300">
      	<input type="button" value="����" onclick="PopItemSellEdit('<%= itemid %>');">
      	</td>
      	<td width="60">��۱��� :</td>
      	<td colspan=2><%= ojaegoitem.FItemList(0).GetDeliveryName %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>��ǰ�ڵ� :</td>
      	<td>10 <b><%= CHKIIF(ojaegoitem.FItemList(0).FItemID>=1000000,Format00(8,ojaegoitem.FItemList(0).FItemID),Format00(6,ojaegoitem.FItemList(0).FItemID)) %></b> <%= itemoption %></td>
      	<td>���ÿ��� : </td>
      	<td colspan=2></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>�귣��ID :</td>
      	<td><%= ojaegoitem.FItemList(0).FMakerid %></td>
      	<td>�Ǹſ��� : </td>
      	<td colspan=2><%= ojaegoitem.FItemList(0).FSellyn %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>��ǰ�� :</td>
      	<td><%= ojaegoitem.FItemList(0).FItemName %></td>
      	<td>��뿩�� : </td>
      	<td colspan=2><%= ojaegoitem.FItemList(0).FIsUsing %></td>
    </tr>
    <% for i=0 to ojaegoitem.FResultCount -1 %>
	    <% if ojaegoitem.FItemList(i).Foptionusing<>"Y" then %>
	    <tr bgcolor="#FFFFFF">
	      	<td><font color="#AAAAAA">�ɼǸ� :</font></td>
	      	<td><font color="#AAAAAA"><%= ojaegoitem.FItemList(i).FItemOptionName %></font></td>
	      	<td><font color="#AAAAAA">�������� : </font></td>
	      	<td><font color="#AAAAAA"><%= ojaegoitem.FItemList(i).FLimitYn %> (<%= ojaegoitem.FItemList(i).GetLimitStr %>)</font></td>
	      	<td>
	      		<%= ojaegoitem.FItemList(i).Foldstockcurrno %> : (OLD)
	      		<%= ojaegoitem.FItemList(i).GetCheckStockNo %> : (NEW)
	      	</td>
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
	      	<td>
	      		<%= ojaegoitem.FItemList(i).Foldstockcurrno %> : (OLD)
	      		<%= ojaegoitem.FItemList(i).GetCheckStockNo %> : (NEW)
	      	</td>
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
		<br><br>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->



<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="20" valign="bottom">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td><b>*���� ���</b>(���ؽð� : ���� ���� 1��)</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->
<%

dim colcount, offtotalstock
dim offtotipno, offtotreno, offtotupcheipno, offtotupchereno, offtotsellno, offtotcurrno
colcount = offstock.FResultCount
dim fromdate, todate

fromdate = "2001-10-10"
todate = Left(now(), 10)

%>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="100">shopID</td>
    	<td width="60">�ŷ�����</td>
    	<td width="60">�԰�<br>(�ٹ�����)</td>
    	<td width="60">��ǰ<br>(�ٹ�����)</td>
    	<td width="60">�԰�<br>(��ü)</td>
    	<td width="60">��ǰ<br>(��ü)</td>
    	<td width="60">���Ǹ�</td>
    	<td width="60" bgcolor="F4F4F4">�ý������</td>
    	<td width="60">����</td>
    	<td width="60">�ҷ�</td>
    	<td width="60" bgcolor="F4F4F4">��ȿ���</td>
    	<td width="60">����</td>
    	<td width="60" bgcolor="F4F4F4">�������</td>
    	<td>���</td>
    </tr>
    <% for i=0 to offstock.FResultcount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= offstock.FItemList(i).FShopid %></td>
    	<td><acronym title="���Ը��� : <%= offstock.FItemList(i).Fdefaultmargin %>&#13���޸��� : <%= offstock.FItemList(i).Fdefaultsuplymargin %>"><font color="<%= offstock.FItemList(i).getChargeDivColor %>"><%= offstock.FItemList(i).GetChargedivName %></font></acronym></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= fromdate %>','<%= todate %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= offstock.FItemList(i).FShopid %>');"><%= offstock.FItemList(i).Fipno %></a></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= fromdate %>','<%= todate %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= offstock.FItemList(i).FShopid %>');"><%= -1 * offstock.FItemList(i).Freno %></a></td>
    	<td><%= offstock.FItemList(i).Fupcheipno %></td>
    	<td><%= -1 * offstock.FItemList(i).Fupchereno %></td>
    	<td><%= -1 * offstock.FItemList(i).Fsellno %></td>
    	<td bgcolor="F4F4F4"><b><%= offstock.FItemList(i).Fcurrno %></b></td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"></td>
    	<td></td>
    	<td bgcolor="F4F4F4"></td>
    	<td></td>
    </tr>
    <%
    if not IsNULL(offstock.FItemList(i).Fipno) then 	offtotipno = offtotipno + offstock.FItemList(i).Fipno
    if not IsNULL(offstock.FItemList(i).Freno) then 	offtotreno = offtotreno + offstock.FItemList(i).Freno
    if not IsNULL(offstock.FItemList(i).Fupcheipno) then 	offtotupcheipno = offtotupcheipno + offstock.FItemList(i).Fupcheipno
    if not IsNULL(offstock.FItemList(i).Fupchereno) then 	offtotupchereno = offtotupchereno + offstock.FItemList(i).Fupchereno
    if not IsNULL(offstock.FItemList(i).Fsellno) then 	offtotsellno = offtotsellno + offstock.FItemList(i).Fsellno
    if not IsNULL(offstock.FItemList(i).Fcurrno) then 	offtotalstock = offtotalstock + offstock.FItemList(i).Fcurrno
    %>
    <% next %>
    <tr align="center" bgcolor="#EEEEEE">
    	<td></td>
    	<td></td>
    	<td><%= offtotipno %></td>
    	<td><%= -1 * offtotreno %></td>
    	<td><%= offtotupcheipno %></td>
    	<td><%= -1 * offtotupchereno %></td>
    	<td><%= offtotsellno %></td>
    	<td bgcolor="F4F4F4"><b><%= offtotalstock %></b></td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"></td>
    	<td></td>
    	<td bgcolor="F4F4F4"></td>
    	<td></td>
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
set osummarystock = Nothing
set offstock = Nothing
%>
<form name=frmrefresh method=post action="dostockrefresh.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="refreshstartdate" value="<%= BasicMonth + "-01" %>">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="<%= itemoption %>">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->