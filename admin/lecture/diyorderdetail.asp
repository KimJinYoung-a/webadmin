<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/lecturecls.asp"-->
<%
dim itemid,odetail
itemid = request("itemid")
set odetail = new CLecture
odetail.FRectItemID = itemid
odetail.GetLectureRegList

dim i
dim totno

totno =0
%>

<table width="820" border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td>�ֹ���ȣ</td>
	<td>����</td>
	<td>����</td>
	<td>���̵�</td>
	<td>����</td>
	<td>��ȭ</td>
	<td>�ڵ���</td>
	<td>�̸���</td>
	<td>�ֹ���</td>
	<td>������</td>
</tr>
<% for i=0 to odetail.FResultCount -1 %>
<%
if Not odetail.FItemList(i).IsCancel then
totno = totno + odetail.FItemList(i).Fitemno
end if
%>
<tr bgcolor="#FFFFFF">
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FOrderserial %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IpkumDivColor %> ><%= odetail.FItemList(i).IpkumDivName %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FBuyName %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FUserID %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).Fitemno %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FBuyPhone %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FBuyHp %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FUserEmail %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FRegdate %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FIpkumDate %></font></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan=4></td>
	<td><%= totno %></td>
	<td colspan=5></td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->