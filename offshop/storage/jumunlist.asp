<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionOffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
' �����ϴµ�
response.end

dim page, shopid
page = request("page")
if page="" then page=1
shopid = session("ssBctId")

dim osheet
set osheet = new COrderSheet
osheet.FCurrPage = page
osheet.Fpagesize=20
osheet.FRectBaljuId = shopid
osheet.GetOrderSheetList


dim i
dim totaljumunsuply, totalfixsuply, totaljumunsellcash
%>
<script language='javascript'>
function PopIpgoSheet(v,itype){
	var popwin;
	popwin = window.open('popshopjumunsheet.asp?idx=' + v + '&itype=' + itype,'shopjumunsheet','width=680,height=600,scrollbars=yes,status=no');
	popwin.focus();
}
</script>
<table width="100%" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
	<td width=80 rowspan="2">�ֹ��ڵ�</td>
	<td width=110 rowspan="2">����ó</td>
	<td width=80>�ֹ���</td>
	<td width=80>������</td>
	<td width=100 rowspan="2">�ֹ�����</td>
	<td width=80 rowspan="2">���ֹ���<br>(�Һ��ڰ�)</td>
	<td width=80 rowspan="2">���ֹ���<br>(���ް�)</td>
	<td width=80 rowspan="2">Ȯ���ݾ�<br>(���ް�)</td>
	<td width=80 rowspan="2">�߼���</td>
	<td width=100 rowspan="2">�����ȣ</td>
	<td width=50 rowspan="2">������</td>
</tr>
<tr bgcolor="#DDDDFF" align=center>
	<td >�԰��û��</td>
	<td >�Ա���</td>
</tr>
<% if osheet.FResultCount >0 then %>
<% for i=0 to osheet.FResultcount-1 %>
<%
totaljumunsellcash = totaljumunsellcash + osheet.FItemList(i).Fjumunsellcash
totaljumunsuply = totaljumunsuply + osheet.FItemList(i).Fjumunsuplycash
totalfixsuply   = totalfixsuply + osheet.FItemList(i).Ftotalsuplycash
%>
<tr bgcolor="#FFFFFF">
	<td align=center rowspan="2"><a href="jumuninputedit.asp?idx=<%= osheet.FItemList(i).Fidx %>"><%= osheet.FItemList(i).Fbaljucode %></a></td>
	<td align=center rowspan="2"><%= osheet.FItemList(i).Ftargetid %><br>(<%= osheet.FItemList(i).Ftargetname %>)</td>
	<td align=center ><%= Left(osheet.FItemList(i).FRegdate,10) %></td>
	<td align=center ><%= Left(osheet.FItemList(i).Fsegumdate,10) %></td>
	<td align=center rowspan="2"><font color="<%= osheet.FItemList(i).GetStateColor %>"><%= osheet.FItemList(i).GetStateName %></font></td>
	<td align=right rowspan="2"><%= FormatNumber(osheet.FItemList(i).Fjumunsellcash,0) %></td>
	<td align=right rowspan="2"><%= FormatNumber(osheet.FItemList(i).Fjumunsuplycash,0) %></td>
	<td align=right rowspan="2"><%= FormatNumber(osheet.FItemList(i).Ftotalsuplycash,0) %></td>
	<td align=center rowspan="2"><%= Left(osheet.FItemList(i).Fbeasongdate,10) %></td>
	<td align=center rowspan="2"><%= Left(osheet.FItemList(i).Fsongjangno,10) %></td>
	<td align=center rowspan="2"><a href="javascript:PopIpgoSheet('<%= osheet.FItemList(i).FIdx %>','2');">��</a>/<a href="javascript:PopIpgoSheet('<%= osheet.FItemList(i).FIdx %>','1');">��</a></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align=center ><%= Left(osheet.FItemList(i).Fscheduledate,10) %></td>
	<td align=center ><%= Left(osheet.FItemList(i).Fipkumdate,10) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td>�Ѱ�</td>
	<td colspan=4></td>
	<td align=right><%= formatNumber(totaljumunsellcash,0) %></td>
	<td align=right><%= formatNumber(totaljumunsuply,0) %></td>
	<td align=right><%= formatNumber(totalfixsuply,0) %></td>
	<td colspan=3></td>
</tr>
<tr bgcolor="#FFFFFF" height=20>
	<td colspan=11 align=center>
	<% if osheet.HasPreScroll then %>
		<a href="?page=<%= osheet.StartScrollPage-1 %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + osheet.StartScrollPage to osheet.FScrollCount + osheet.StartScrollPage - 1 %>
		<% if i>osheet.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if osheet.HasNextScroll then %>
		<a href="?page=<%= i %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=11 align=center>[ �˻������ �����ϴ�. ]</td>
</tr>
<% end if %>
</table>
<%
set osheet = Nothing
%>


<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->