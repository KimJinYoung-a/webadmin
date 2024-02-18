<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��� Ȯ�� ������
' History : 2007.09.11 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/new_itemcls.asp"-->
<!-- #include virtual="/lib/classes/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/classes/offshop_dailystock.asp"-->
<!-- #include virtual="/lib/classes/auction/auctionclass.asp"-->

<%
const C_STOCK_DAY=7
dim itemgubun, itemid, itemoption 		'��������
itemgubun = request("itemgubun")		'��ǰ������ �޾ƿ´�
itemid = request("itemid")				'��ǰid �޾ƿ�
itemoption = request("itemoption")		'��ǰ�ɼ��ڵ� �޾ƿ�
	if itemgubun="" then 				'��ǰ������ �����̶��
		itemgubun="10"					'�⺻�� 10 �Է�
	end if
	if itemoption="" then 				'��ǰ�ɼ��ڵ尡 �����̸�
		itemoption="0000"				'�⺻�� 0000 �Է�
	end if

dim oitem
set oitem = new CItemInfo				'������ Ŭ���� �ְ�
oitem.FRectItemID = itemid				'��ǰid�� �ְ�
	if itemid<>"" then					'��ǰid�� �����̸�
		oitem.GetOneItemInfo
	end if

dim oitemoption							'�����ۿɼǺκ�
set oitemoption = new CItemOption		'Ŭ���� �ְ�
oitemoption.FRectItemID = itemid
	if itemid<>"" then
		oitemoption.GetItemOptionInfo
	end if

	if (oitemoption.FResultCount<1) then	'��ǰ�ɼ��ڵ尡 1���� �۴ٸ�
		itemoption = "0000"					'�⺻�� 0000 �ְ�
	end if

dim offstock			'������������ľ�
set offstock = new COffShopDailyStock		'Ŭ�����ְ�
offstock.FRectItemGubun = itemgubun
offstock.FRectItemid = itemid
offstock.FRectItemoption = itemoption
	if itemid<>"" then
			if oitem.FResultCount>0 then
				offstock.FRectMakerid = oitem.FOneItem.FMakerid
			end if
		offstock.GetCurrentAllShopItemStock
	end if

dim BasicMonth
BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)

dim osummarystock										'�¶�������ľ�
set osummarystock = new CSummaryItemStock				'Ŭ���� �ְ�
osummarystock.FRectStartDate = BasicMonth + "-01"
osummarystock.FRectItemGubun = itemgubun
osummarystock.FRectItemID =  itemid
osummarystock.FRectItemOption =  itemoption
	if itemid<>"" then
		osummarystock.GetCurrentItemStock
		osummarystock.GetDaily_Logisstock_Summary
	end if

dim i,menupos
	menupos = request("menupos")
%>

<script language="javascript">

	function addfrm() {
		jaegoaddfrm.target= "view";
		jaegoaddfrm.submit();
	}

</script>

<!-- ǥ �˻��κ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>��� Ȯ��</strong> / ���ϻ�ǰ�� ��ϵ��� �ʽ��ϴ�.</font>
			</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<p align="right"><a href="/admin/auction/auction_add_re.asp">�ϰ����</a></p>
			��ǰ�ڵ�: <input type=text name=itemid value="<%= itemid %>" size=9 maxlength=9>
			&nbsp;&nbsp;
			<input type=button value="�˻�" onclick="document.frm.submit();">
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</form>
<form name="jaegoaddfrm" method=post action="/admin/auction/auction_process.asp">
<input type="hidden" name="fmode" value="item_add">
</table>
<!-- ǥ �˻��κ� ��-->

<!-- ��ǰ ���� ����-->
<% if oitem.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=<%= 6 + oitemoption.FResultCount -1 %> width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
      	<td width="60">��ǰ����</td>
      	<td width="300">
      		<% if itemgubun = 10 then %>
				�¶��λ�ǰ
			<% elseif itemgubun = 90 then %>
				�������λ�ǰ
			<% elseif itemgubun = 70 then %>
				�Ҹ�ǰ
			<% end if %>
      	</td>
      	<td width="60">��۱��� :</td>
      	<td colspan=2><%= oitem.FOneItem.GetDeliveryName %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>��ǰ�ڵ� :</td>
      	<td><%= Format00(5,oitem.FOneItem.FItemID) %></td>
      	<td>���ÿ��� : </td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FDispyn) %>"><%= oitem.FOneItem.FDispyn %></font></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>�귣��ID :</td>
      	<td><%= oitem.FOneItem.FMakerid %></td>
      	<td>�Ǹſ��� : </td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FSellyn) %>"><%= oitem.FOneItem.FSellyn %></font></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>��ǰ�� :</td>
      	<td><%= oitem.FOneItem.FItemName %></td>
      	<td>��뿩�� : </td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FIsUsing) %>"><%= oitem.FOneItem.FIsUsing %></font></td>
    </tr>

    <% if oitemoption.FResultCount>1 then %>

		<!-- �ɼ��� �ִ°�� -->
	    <% for i=0 to oitemoption.FResultCount -1 %>
		    <% if oitemoption.FITemList(i).FOptIsUsing<>"Y" then %>
		    <tr bgcolor="#FFFFFF">
		      	<td><font color="#AAAAAA">�ɼǸ� :</font></td>
		      	<td><font color="#AAAAAA"><%= oitemoption.FITemList(i).foptionname %></font></td>
		      	<td><font color="#AAAAAA">�������� : </font></td>
		      	<td><font color="#AAAAAA"><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</font></td>
		      	<td>���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
		    </tr>
		    <% else %>

		    <% if oitemoption.FITemList(i).Fitemoption=itemoption then %>
		    <tr bgcolor="#EEEEEE">
		    <% else %>
		    <tr bgcolor="#FFFFFF">
		    <% end if %>
		      	<td>�ɼǸ� :</td>
		      	<td><%= oitemoption.FITemList(i).foptionname %></td>
		      	<td>�������� : </td>
		      	<td><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
		      	<td>���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo%></b>)</td>
		    </tr>
		    <% end if %>
	    <% next %>
    <% else %>
    	<tr bgcolor="#FFFFFF">
	      	<td>�ɼ��ڵ� :</td>
	      	<td>-<input type="hidden" value="0000" name="itemoption"></td>
	      	<td>�������� : </td>
	      	<td><font color="<%= ynColor(oitem.FOneItem.Flimityn) %>"><%= oitem.FOneItem.Flimityn %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
	      	<td>���� ����� (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)</td>
	    </tr>
    <% end if %>
</table>
<!-- ��ǰ ���� ��-->
</table>
<% dim oip
	set oip = new Cauctionlist        	'Ŭ���� ����
	oip.Frectitemid = itemid
	oip.fwritelist()					'Ŭ������ ����
%>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
<tr bgcolor="#FFFFFF">
<td>���� ī�װ� :</td>
<td><input type="text" name="auction_cate_code" value="10060500"></td>
<td>������ :</td>
<td><input type="text" name="wonsanji" value="�ѱ�"> ex) �ѱ� , ����</td>
<td>
<input type="hidden" name="refreshstartdate" value="<%= BasicMonth + "-01" %>">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>,">
<input type="hidden" name="makerid" value="<%= oitem.FOneItem.FMakerid %>">
<input type="hidden" name="imagesrc" value="<%= oitem.FOneItem.FListImage %>">
<input type="hidden" name="itemname" value="<%= oitem.FOneItem.FItemName %>">
<input type="button" value="����" onclick=addfrm();></td>
</tr></form>
<tr bgcolor="#FFFFFF">
<td>��ǰ���� :</td>
<td colspan="4"><textarea name="ten_itemcontent" cols="80" rows="30"><%= oip.flist(0).fitemcontent %></textarea></td>
</tr>
</table>

<% else %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
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
set oitemoption = Nothing
set oitem = Nothing
set osummarystock = Nothing
%>

<iframe frameboarder=0 height=0 width=0 name="view" id="view"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->