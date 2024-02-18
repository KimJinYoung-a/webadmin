<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/offReportClass.asp" -->
<%
	Dim cStatistic, cStatistic2, i, yyyy1, mm1
	yyyy1    = requestCheckVar(request("yyyy1"),4)
	mm1    = requestCheckVar(request("mm1"),2)
	Set cStatistic = New COffReport
	Set cStatistic2 = New COffReport
	cStatistic.FRectYYYYMM = yyyy1 & "-" & TwoNumber(mm1)
	if yyyy1 <> "" then
	cStatistic.GetSoldoutCancelOrderSet
	cStatistic.GetSoldoutCancelOrderInfo1
	cStatistic.GetSoldoutCancelOrderInfo2
	cStatistic2.FRectYYYYMM = yyyy1 & "-" & TwoNumber(mm1)
	cStatistic2.GetSoldoutCancelOrderInfo3
	end if
	yyyy1=NullFillWith(request("yyyy1"),Year(now))
	mm1=NullFillWith(request("mm1"),Month(now))
%>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<font color="#CC3333">��/�� :</font> <% DrawYMBox yyyy1,mm1 %>
		</td>

		<td bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� --><br>
<table width="10%" align="left" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">��� �� �Ǽ�</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= cStatistic.FTotalCount %></td>
</tr>
</table>
<table width="20%" align="left" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td colspan="2" align="center">��ҰǱ���</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">����</td>
	<td align="center">�Ǽ�</td>
</tr>
<% for i=0 to cStatistic.FResultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= cStatistic.FItemList(i).Fcomm_name %></td>
	<td align="center"><%= cStatistic.FItemList(i).Ftotalcnt %></td>
</tr>
<% next %>
</table>
<table width="20%" align="left" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td colspan="2" align="center">ǰ���Ǳ���</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">����</td>
	<td align="center">�Ǽ�</td>
</tr>
<% for i=0 to cStatistic2.FResultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><% if cStatistic2.FItemList(i).FallianceYN="Y" then %>����ǰ����<% else %>10x10ǰ����<% end if %></td>
	<td align="center"><%= cStatistic2.FItemList(i).Ftotalcnt %></td>
</tr>
<% next %>
</table>
<%
Set cStatistic = Nothing
Set cStatistic2 = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->