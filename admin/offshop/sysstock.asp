<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ���
' History : �̻� ����
'			2017.04.11 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->

<%
dim makerid, shopid, availstock, research
makerid = requestCheckVar(request("makerid"),32)
shopid = requestCheckVar(request("shopid"),32)
availstock = requestCheckVar(request("availstock"),32)
research = requestCheckVar(request("research"),2)

if (research="") and (availstock="") then availstock="on"

dim offstock
set offstock = new COffShopDailyStock
offstock.FRectShopId = shopid
offstock.FRectMakerid = makerid
offstock.FRecAvailStock = availstock

if (makerid<>"") and (shopid<>"") then
	offstock.GetCurrentSysStock
end if

dim i, iptot,retot,upiptot,upretot,selltot,currtot
dim iptotsum, retotsum, upiptotsum, upretotsum , selltotsum, currtotsum
%>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr>
	<td class="a" >
		�� : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
		��ü:<% drawSelectBoxDesignerwithName "makerid",makerid  %> &nbsp;&nbsp;
		<!--
		<input type=checkbox name="availstock" <% if availstock="on" then response.write "checked" %> >��ȿ����˻�
		-->
	</td>
	<td class="a" align="right">
		<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	</td>
</tr>
</form>
</table>
<br>
* ���� �԰�/��ǰ �� ��ü �԰�/��ǰ�� ���� ��������� �־�� (�԰�ó�� �Ǿ��)  ���˴ϴ�.

<table width="100%" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
	<td width="86">���ڵ�</td>
	<td width="100">��ǰ��</td>
	<td width="80">�ɼǸ�</td>
	<td width="60">�ǸŰ�</td>
	<!--
	<td width="60">�¶��ΰ���</td>
	-->
	<td width="50">����<br>�԰�</td>
	<td width="50">����<br>��ǰ</td>
	<td width="50">��ü<br>�԰�</td>
	<td width="50">��ü<br>��ǰ</td>
	<td width="50">�Ǹŷ�</td>
	<td width="50">�������</td>
</tr>
<% for i=0 to offstock.FresultCount-1 %>
<%
	iptot = iptot + offstock.FItemList(i).Fipno
	retot = retot + offstock.FItemList(i).Freno
	upiptot = upiptot + offstock.FItemList(i).Fupcheipno
	upretot = upretot +  offstock.FItemList(i).Fupchereno
	selltot = selltot + offstock.FItemList(i).Fsellno
	currtot = currtot + offstock.FItemList(i).Fcurrno


	iptotsum = iptotsum + offstock.FItemList(i).Fipno * offstock.FItemList(i).Fshopitemprice
	retotsum = retotsum + offstock.FItemList(i).Freno * offstock.FItemList(i).Fshopitemprice
	upiptotsum = upiptotsum + offstock.FItemList(i).Fupcheipno * offstock.FItemList(i).Fshopitemprice
	upretotsum = upretotsum + offstock.FItemList(i).Fupchereno * offstock.FItemList(i).Fshopitemprice
	selltotsum = selltotsum + offstock.FItemList(i).Fsellno * offstock.FItemList(i).Fshopitemprice
	currtotsum = currtotsum + offstock.FItemList(i).Fcurrno * offstock.FItemList(i).Fshopitemprice

%>
<tr bgcolor="#FFFFFF">
	<td><%= offstock.FItemList(i).GetBarCode %></td>
	<td><%= offstock.FItemList(i).FItemName %></td>
	<td><%= offstock.FItemList(i).FItemOptionName %></td>
	<td align=right>
	<%= formatNumber(offstock.FItemList(i).Fshopitemprice,0) %>
	</td>
	<!--
	<td align=right><%= formatNumber(offstock.FItemList(i).Fonlinesellcash,0) %></td>
	-->
	<td align="center"><%= offstock.FItemList(i).Fipno  %></td>
	<td align="center"><%= offstock.FItemList(i).Freno %></td>
	<td align="center"><%= offstock.FItemList(i).Fupcheipno %></td>
	<td align="center"><%=  offstock.FItemList(i).Fupchereno %></td>
	<td align="center"><%= offstock.FItemList(i).Fsellno %></td>
	<% if offstock.FItemList(i).Fcurrno<1 then %>
	<td align="center"><font color="red"><b><%= offstock.FItemList(i).Fcurrno %></font></b></td>
	<% else %>
	<td align="center"><%= offstock.FItemList(i).Fcurrno %></td>
	<% end if %>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="3">total</td>
	<td align="center"></td>
	<td align="center"><%= iptot %></td>
	<td align="center"><%= retot %></td>
	<td align="center"><%= upiptot %></td>
	<td align="center"><%= upretot %></td>
	<td align="center"><%= selltot %></td>
	<td align="center"><%= currtot %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="3">�������԰��</td>
	<td colspan="7" ><%= FormatNumber(iptotsum,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="3">�����ѹ�ǰ��</td>
	<td colspan="7"><%= FormatNumber(retotsum,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="3">��ü���԰��</td>
	<td colspan="7"><%= FormatNumber(upiptotsum,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="3">��ü�ѹ�ǰ��</td>
	<td colspan="7"><%= FormatNumber(upretotsum,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="3">�Ѹ����</td>
	<td colspan="7"><%= FormatNumber(selltotsum,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="3">������</td>
	<td colspan="7"><%= FormatNumber(currtotsum,0) %></td>
</tr>
</table>
<%
set offstock = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->