<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->

<%
dim oip, i
set oip = new organizerCls
	oip.FOrgStoryList

%>

<!-- ����Ʈ ���� -->
<table border="0" cellpadding="0" cellspacing="0">
<tr>
	<td><input type="button" onClick="location.href='storyview.asp';" value="���۾���"></td>
</tr>
</table>
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<% IF oip.ftotalcount>0 Then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= oip.ftotalcount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td nowrap> ��ȣ</td>
		<td nowrap> ���� </td>
		<td nowrap> �ۼ��� </td>
		<td nowrap> ���� </td>
	</tr>

	<% For i =0 To  oip.ftotalcount-1 %>
	<tr bgcolor="#FFFFFF">
		<td nowrap align="center"><%= oip.FItemList(i).FOW_IDX %> </td>
		<td nowrap><%= Left(oip.FItemList(i).FOW_TITLE,50) %>... </td>
		<td nowrap align="center"><%=oip.FItemList(i).FOW_REGDATE %></td>
		<td nowrap align="center"><input type="button" onClick="location.href='storyview.asp?idx=<%=oip.FItemList(i).FOW_IDX %>';" value="����"></td>
	</tr>
	<% Next %>
<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
<% End IF %>

</table>
<!-- ����Ʈ �� -->

<% set oip = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->