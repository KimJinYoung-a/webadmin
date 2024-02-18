<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ���̳ʽ� ���
' History : �̻� ����
'			2017.04.13 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<%
dim shopid,jaegono, makerid, page
shopid = requestCheckVar(request("shopid"),32)
jaegono = requestCheckVar(request("jaegono"),10)
makerid = requestCheckVar(request("makerid"),32)
page = requestCheckVar(request("page"),10)

if (jaegono="") then jaegono=1
if (page="") then page=1

dim offstock
set offstock = new COffShopDailyStock
offstock.FCurrPage = page
offstock.FPageSize = 100
offstock.FRectMinusNo = jaegono
offstock.FRectMakerid = makerid
offstock.FRectShopId = shopid

if (shopid<>"") then
	offstock.GetCurrentStockMinusList
end if

dim i, iptot,retot,selltot,currtot
%>
<script language='javascript'>
function NextPage(p){
	document.frm.page.value = p;
	document.frm.submit();
}
</script>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<tr>
		<td class="a" >
			�� : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
			��ü:<% drawSelectBoxDesignerwithName "makerid",makerid  %>
			<br>
			�������
			<input type="text" name="jaegono" value="<%= jaegono %>" size="3" maxlength="4">
			�� �̸�
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width=800 class=a>
<tr>
	<td align=right>�� <%= offstock.FTotalCount %> �� <%= page %>/<%= offstock.FtotalPage %> page</td>
</tr>
</table>
<br>
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
    <td width="50">�̹���</td>
	<td width="86">���ڵ�</td>
	<td width="100">��ǰ��</td>
	<td width="80">�ɼǸ�</td>
	<td width="80">����<br>�ǻ���</td>
	<td width="50">����<br>�ǻ�</td>
	<td width="50">�԰�</td>
	<td width="50">��ǰ</td>
	<td width="50">�Ǹŷ�</td>
	<td width="50">�������</td>
</tr>
<% if (shopid="") then %>
<tr bgcolor="#FFFFFF">
	<td colspan=10 align=center><font color=red>���� ������ �ּ���.</font></td>
</tr>
<% else %>
<% for i=0 to offstock.FresultCount-1 %>
<%
	iptot = iptot + offstock.FItemList(i).Fipno + offstock.FItemList(i).Fupcheipno
	retot = retot + offstock.FItemList(i).Freno + offstock.FItemList(i).Fupchereno
	selltot = selltot + offstock.FItemList(i).Fsellno
	currtot = currtot + offstock.FItemList(i).Fcurrno
%>
<tr bgcolor="#FFFFFF">
	<td><img src="<%= offstock.FItemList(i).Fimgsmall %>" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" width=50 height=50></td>
	<td><%= offstock.FItemList(i).GetBarCode %></td>
	<td><%= offstock.FItemList(i).FItemName %></td>
	<td><%= offstock.FItemList(i).FItemOptionName %></td>
	<td align="center"><%= offstock.FItemList(i).Flastrealdate %></td>
	<td align="center"><%= offstock.FItemList(i).Flastrealno %></td>
	<td align="center"><%= offstock.FItemList(i).Fipno + offstock.FItemList(i).Fupcheipno %></td>
	<td align="center"><%= offstock.FItemList(i).Freno + offstock.FItemList(i).Fupchereno %></td>
	<td align="center"><%= offstock.FItemList(i).Fsellno %></td>
	<% if offstock.FItemList(i).Fcurrno<1 then %>
	<td align="center"><font color="red"><b><%= offstock.FItemList(i).Fcurrno %></font></b></td>
	<% else %>
	<td align="center"><%= offstock.FItemList(i).Fcurrno %></td>
	<% end if %>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="5">total</td>
	<td align="center"></td>
	<td align="center"><%= iptot %></td>
	<td align="center"><%= retot %></td>
	<td align="center"><%= selltot %></td>
	<td align="center"><%= currtot %></td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" align="center">
	<% if offstock.HasPreScroll then %>
		<a href="javascript:NextPage('<%= offstock.StartScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + offstock.StartScrollPage to offstock.FScrollCount + offstock.StartScrollPage - 1 %>
		<% if i>offstock.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if offstock.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set offstock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->