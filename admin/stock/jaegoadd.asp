<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����ľ� ���� ����������
' History : 2007.07.13 �ѿ�� ����
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

dim i , submitview
	submitview = request("submitview")
%>

<script language="javascript">

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function addfrm(jaegoaddfrm) {
jaegoaddfrm.submit();
}
</script>

<!--ǥ ������-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>����ľ� �������</strong> / ���ϻ�ǰ�� ����ľ����ϰ�� ��ϵ��� �ʽ��ϴ�.
			</td>

		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>
	</tr>
</table>
<!--ǥ ��峡-->

<!-- ǥ �˻��κ� ����-->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method=get>
		<tr bgcolor="#FFFFFF" valign="top">
	        <td background="/images/tbl_blue_round_04.gif" width="1%" bgcolor="F4F4F4"></td>
	        <td width="54%" bgcolor="F4F4F4">
	       		��ǰ�ڵ�: <input type=text name=itemid value="<%= itemid %>" size=9 maxlength=9>
	        	&nbsp;&nbsp;
	        	<input type=button value="�˻�" onclick="document.frm.submit();">
	        </td>
	        <td valign="top" align="right" width="40%" bgcolor="F4F4F4">
	      	</td>
	        <td background="/images/tbl_blue_round_05.gif" bgcolor="F4F4F4" width="1%"></td>
	    </tr>
    </form>
    <form name="jaegoaddfrm" method=post action="jaegocheck_process.asp">
</table>
<!-- ǥ �˻��κ� ��-->

<!-- ��ǰ ���� ����-->
<% if oitem.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=<%= 7 + oitemoption.FResultCount -1 %> width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
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
      	<td colspan=2>
      		<% if oitem.FOneItem.Fdanjongyn="Y" then %>
			<font color="#33CC33">����</font>
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
		    <% if oitemoption.FITemList(i).FOptIsUsing<>"Y" then %>
		    <tr bgcolor="#FFFFFF">
		      	<td><font color="#AAAAAA">�ɼ��ڵ� :</font></td>
		      	<td><font color="#AAAAAA"><%= oitemoption.FITemList(i).Fitemoption %></font></td>
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
		      	<td>�ɼ��ڵ� :</td>
		      	<td><%= oitemoption.FITemList(i).Fitemoption %></td>
		      	<td>�������� : </td>
		      	<td><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
		      	<td>���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
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

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<tr height="20" valign="bottom">
        	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        	<td><b>* �����ľǵ����</b></td>
       		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ �߰��� ��-->


<!--��������Ȳ����-->
<table width="100%" align="center" cellpadding="2" cellspacing="1" bgcolor="#BABABA" class="a">
 	<tr align="center" bgcolor="#DDDDFF">
    		<td width="50">��<br>�԰�/��ǰ</td>
    		<td width="50">ON��<br>�Ǹ�/��ǰ</td>
			<td width="50">OFF��<br>���/��ǰ</td>
			<td width="50">��Ÿ<br>���/��ǰ</td>
			<td width="50" bgcolor="F4F4F4">�ý���<br>�����</td>
			<td width="50">�Ѻҷ�</td>
	      		<td width="50" bgcolor="F4F4F4">�ý���<br>��ȿ���</td>
			<td width="50">�ѽǻ�<br>����</td>
			<td width="50" bgcolor="F4F4F4">�ǻ�<br>���</td>
			<td width="50">ON<br>��ǰ�غ�</td>
			<td width="50">OFF<br>��ǰ�غ�</td>
			<td width="50" bgcolor="F4F4F4">����ľ�<br>���</td>
			<td width="50">ON<br>�����Ϸ�</td>
			<td width="50">ON<br>�ֹ�����</td>
			<td width="50">OFF<br>�ֹ�����</td>
			<td bgcolor="F4F4F4">����<br>���</td>

	</tr>
	<tr bgcolor="#FFFFFF" height="25" align=center>
    		<td><%= osummarystock.FOneItem.Ftotipgono %></td>
    		<td><%= -1*osummarystock.FOneItem.Ftotsellno %></td>
    		<td><%= osummarystock.FOneItem.Foffchulgono + osummarystock.FOneItem.Foffrechulgono %></td>
    		<td><%= osummarystock.FOneItem.Fetcchulgono + osummarystock.FOneItem.Fetcrechulgono %></td>
    		<td bgcolor="F4F4F4"><b><%= osummarystock.FOneItem.Ftotsysstock %></b></td>
    		<td><%= osummarystock.FOneItem.Ferrbaditemno %></td>
    		<td bgcolor="F4F4F4"><b><%= osummarystock.FOneItem.Favailsysstock %></b></td>
    		<td><%= osummarystock.FOneItem.Ferrrealcheckno %></td>
    		<td bgcolor="F4F4F4"><b><%= osummarystock.FOneItem.Frealstock %></td>
    		<td><%= osummarystock.FOneItem.Fipkumdiv5 %></td>
    		<td><%= osummarystock.FOneItem.Foffconfirmno %></td>
    		<td bgcolor="F4F4F4"><b><%= osummarystock.FOneItem.GetCheckStockNo %></b></td>
    		<td><%= osummarystock.FOneItem.Fipkumdiv4 %></td>
    		<td><%= osummarystock.FOneItem.Fipkumdiv2 %></td>
    		<td><%= osummarystock.FOneItem.Foffjupno %></td>
    		<td bgcolor="F4F4F4"><b><%= osummarystock.FOneItem.GetMaystock %></b></td>
	</tr>
    <tr bgcolor="#FFFFFF" height="25" align=center>
    	<td colspan=9><input type="hidden" name="jaego" value="<%= osummarystock.FOneItem.GetCheckStockNo %>">
    	<input type="button" value="����ľ������ϱ�" onclick=addfrm(jaegoaddfrm);></td>
    	<td colspan="2"><%= osummarystock.FOneItem.Fipkumdiv5 + osummarystock.FOneItem.Foffconfirmno %></td>
    	<td></td>
    	<td colspan="3"><%= osummarystock.FOneItem.Fipkumdiv4 + osummarystock.FOneItem.Fipkumdiv2 + osummarystock.FOneItem.Foffjupno %></td>
    	<td></td>
    </tr>
<input type="hidden" name="refreshstartdate" value="<%= BasicMonth + "-01" %>">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="makerid" value="<%= oitem.FOneItem.FMakerid %>">
<input type="hidden" name="imagesrc" value="<%= oitem.FOneItem.FListImage %>">
<input type="hidden" name="itemname" value="<%= oitem.FOneItem.FItemName %>">
</form>
</table>
<!--��������Ȳ��-->


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
        <td valign="bottom" align="right"><input type="button" value="�ݱ�" onclick="javascript:window.close();"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<% if submitview = "yes" then %>
<script language="javascript">
jaegoaddfrm.submit();
</script>
<% end if %>
<%
set oitemoption = Nothing
set oitem = Nothing
set osummarystock = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->